/* eslint-disable @typescript-eslint/no-explicit-any */
import "./App.css";
import { useEffect, useRef, useState } from "react";
import {
  createUniver,
  ICommandService,
  LocaleType,
  mergeLocales,
  Univer,
} from "@univerjs/presets";

// shadcn components
import { Button } from "@/components/ui/button";
import { cn } from "@/lib/utils";

// presets
import { UniverSheetsHyperLinkPreset } from "@univerjs/preset-sheets-hyper-link";
import { UniverSheetsCorePreset } from "@univerjs/preset-sheets-core";
import { UniverSheetsDataValidationPreset } from "@univerjs/preset-sheets-data-validation";
import { UniverSheetsDrawingPreset } from "@univerjs/preset-sheets-drawing";
import { UniverSheetsFindReplacePreset } from "@univerjs/preset-sheets-find-replace";

// locales
import UniverPresetSheetsCoreEnUS from "@univerjs/preset-sheets-core/locales/en-US";
import UniverPresetSheetsHyperLinkEnUS from "@univerjs/preset-sheets-hyper-link/locales/en-US";
import UniverPresetSheetsDataValidationEnUS from "@univerjs/preset-sheets-data-validation/locales/en-US";
import UniverPresetSheetsDrawingEnUS from "@univerjs/preset-sheets-drawing/locales/en-US";
import sheetsFindReplaceEnUS from "@univerjs/preset-sheets-find-replace/locales/en-US";

// css
import "@univerjs/preset-sheets-core/lib/index.css";
import "@univerjs/preset-sheets-hyper-link/lib/index.css";
import "@univerjs/preset-sheets-data-validation/lib/index.css";
import "@univerjs/preset-sheets-drawing/lib/index.css";
import "@univerjs/preset-sheets-find-replace/lib/index.css";

// types
import type { FWorkbook } from "@univerjs/sheets/lib/types/facade/f-workbook.js";
import type { Subscription } from "rxjs";

// initial data
import { WORKBOOK_DATA } from "./WORKBOOK_DATA";
import { throttle } from "./lib/utils";

interface CreateUniverResult {
  univerAPI: {
    createWorkbook: (data: unknown) => void;
    getActiveWorkbook?: () => FWorkbook | string | undefined;
    getWorkbook?: (id: string) => FWorkbook | undefined;
    dispose: () => void;
    addEvent?: (event: any, handler: (event: any) => void) => void;
    removeEvent?: (event: any, handler: (event: any) => void) => void;
    Event?: any;
    Enum?: any;
  };
  univer: Univer;
}

interface CommandExecutedEvent {
  id: string;
  type?: string;
  params?: Record<string, unknown>;
}

interface ExtendedCommandService {
  commandExecuted?: {
    subscribe: (
      callback: (event: CommandExecutedEvent) => void
    ) => Subscription;
  };
}

interface ICellEventParam {
  worksheet: any;
  workbook: any;
  row: number;
  column: number;
  event?: Event;
}

function App() {
  const containerRef = useRef<HTMLDivElement>(null);
  const univerRef = useRef<Univer | null>(null);
  const univerAPIRef = useRef<CreateUniverResult["univerAPI"] | null>(null);

  const [hoveredCell, setHoveredCell] = useState<string | null>(null);
  const [isDragging, setIsDragging] = useState(false);
  const isDraggingRef = useRef(false);

  useEffect(() => {
    isDraggingRef.current = isDragging;
  }, [isDragging]);

  const [dragPos, setDragPos] = useState<{ x: number; y: number } | null>(null);
  const [notification, setNotification] = useState<string | null>(null);

  const draggedDataRef = useRef<{ files?: File[]; urls?: string[] } | null>(
    null
  );
  const insertFeedbackRef = useRef<Map<string, number>>(new Map());

  const isDark =
    typeof document !== "undefined" &&
    document.documentElement.classList.contains("dark");

  const showNotification = (message: string) => {
    setNotification(message);
    setTimeout(() => setNotification(null), 3000);
  };

  // Export functionality
  const handleExport = () => {
    const active = univerAPIRef.current?.getActiveWorkbook?.();
    const wb: FWorkbook | undefined =
      active && typeof active !== "string"
        ? active
        : typeof active === "string"
        ? univerAPIRef.current?.getWorkbook?.(active)
        : undefined;

    if (!wb) {
      showNotification("No workbook available to export");
      return;
    }

    let snapshot: unknown;
    if (typeof wb.getSnapshot === "function") snapshot = wb.getSnapshot();
    else if (
      typeof (wb as FWorkbook & { serialize: () => unknown }).serialize ===
      "function"
    )
      snapshot = (wb as FWorkbook & { serialize: () => unknown }).serialize();
    else if (typeof wb.save === "function") snapshot = wb.save();

    if (!snapshot) {
      showNotification("Unable to export workbook data");
      return;
    }

    const jsonString = JSON.stringify(snapshot, null, 2);
    const blob = new Blob([jsonString], { type: "application/json" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = `univer-workbook-${Date.now()}.json`;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);

    showNotification("Workbook exported successfully!");
  };

  // Import functionality - FIXED to prevent render conflicts
  const handleImport = () => {
    const input = document.createElement("input");
    input.type = "file";
    input.accept = ".json";

    input.onchange = (e: Event) => {
      const file = (e.target as HTMLInputElement).files?.[0];
      if (!file) return;

      const reader = new FileReader();
      reader.onload = (event) => {
        try {
          const jsonData = JSON.parse(event.target?.result as string);

          // Clean up old instance
          if (univerAPIRef.current?.dispose) {
            univerAPIRef.current.dispose();
          }
          if (univerRef.current) {
            univerRef.current = null;
          }

          // Small delay to ensure cleanup
          setTimeout(() => {
            initializeUniver(jsonData);
            showNotification("Workbook imported successfully!");
          }, 100);
        } catch (error) {
          console.error("Error importing workbook:", error);
          showNotification("Error importing workbook. Invalid file format.");
        }
      };

      reader.onerror = () => {
        showNotification("Error reading file");
      };

      reader.readAsText(file);
    };

    input.click();
  };

  // Prevent browser navigation on file drop
  useEffect(() => {
    const preventNav = (e: DragEvent) => {
      e.preventDefault();
    };
    window.addEventListener("dragover", preventNav, { passive: false });
    window.addEventListener("drop", preventNav, { passive: false });
    return () => {
      window.removeEventListener("dragover", preventNav);
      window.removeEventListener("drop", preventNav);
    };
  }, []);

  const extractImageUrls = (html: string): string[] => {
    const parser = new DOMParser();
    const doc = parser.parseFromString(html, "text/html");
    return Array.from(doc.querySelectorAll("img"))
      .map((img) => img.src)
      .filter(
        (src) => src && (src.startsWith("http") || src.startsWith("data:image"))
      );
  };

  const collectDragData = (dt: DataTransfer | null | undefined) => {
    const out: { files: File[]; urls: string[] } = { files: [], urls: [] };
    if (!dt) return out;

    try {
      if (dt.files && dt.files.length > 0) {
        out.files = Array.from(dt.files).filter((f) =>
          f.type.startsWith("image/")
        );
      }
    } catch {
      // Ignore
    }

    try {
      const html = dt.getData("text/html");
      if (html) out.urls.push(...extractImageUrls(html));
    } catch {
      // Ignore
    }

    try {
      const uriList = dt.getData("text/uri-list");
      if (uriList) {
        out.urls.push(
          ...uriList
            .split("\n")
            .map((s) => s.trim())
            .filter((s) => s && !s.startsWith("#"))
        );
      }
    } catch {
      // Ignore
    }

    try {
      const plain = dt.getData("text/plain");
      if (
        plain &&
        (plain.startsWith("http") || plain.startsWith("data:image"))
      ) {
        out.urls.push(plain.trim());
      }
    } catch {
      // Ignore
    }

    out.urls = Array.from(new Set(out.urls));
    return out;
  };

  const insertImageToCell = async (file: File, cellRef: string) => {
    return new Promise<void>((resolve) => {
      const reader = new FileReader();
      reader.onload = async (e) => {
        const dataUrl = e.target?.result as string;
        try {
          const workbook = univerAPIRef.current?.getActiveWorkbook?.();
          if (!workbook || typeof workbook === "string") {
            resolve();
            return;
          }
          const sheet = (workbook as any).getActiveSheet();
          if (!sheet) {
            resolve();
            return;
          }
          const range = sheet.getRange(cellRef);
          if (range) {
            await range.insertCellImageAsync(dataUrl);
            console.log(`Image inserted to cell ${cellRef}`);

            // Visual feedback without custom renderer
            insertFeedbackRef.current.set(cellRef, Date.now());
            setTimeout(() => {
              insertFeedbackRef.current.delete(cellRef);
            }, 1200);
          }
        } catch (error) {
          console.error("Error inserting image:", error);
        }
        resolve();
      };
      reader.onerror = () => {
        console.error("Error reading file");
        resolve();
      };
      reader.readAsDataURL(file);
    });
  };

  const insertImageUrlToCell = async (url: string, cellRef: string) => {
    try {
      const workbook = univerAPIRef.current?.getActiveWorkbook?.();
      if (!workbook || typeof workbook === "string") return;
      const sheet = (workbook as any).getActiveSheet();
      if (!sheet) return;
      const range = sheet.getRange(cellRef);
      if (range) {
        await range.insertCellImageAsync(url);
        console.log(`Image from URL inserted to cell ${cellRef}`);

        // Visual feedback
        insertFeedbackRef.current.set(cellRef, Date.now());
        setTimeout(() => {
          insertFeedbackRef.current.delete(cellRef);
        }, 1200);
      }
    } catch (error) {
      console.error("Error inserting image from URL:", error);
    }
  };

  const initializeUniver = (initialData: unknown = WORKBOOK_DATA) => {
    const { univerAPI, univer } = createUniver({
      darkMode: false,
      locale: LocaleType.EN_US,
      locales: {
        [LocaleType.EN_US]: mergeLocales(
          UniverPresetSheetsCoreEnUS,
          UniverPresetSheetsHyperLinkEnUS,
          UniverPresetSheetsDataValidationEnUS,
          UniverPresetSheetsDrawingEnUS,
          sheetsFindReplaceEnUS
        ),
      },
      presets: [
        UniverSheetsCorePreset({
          container: containerRef.current || undefined,
          contextMenu: true,
        }),
        UniverSheetsHyperLinkPreset({
          urlHandler: {
            navigateToOtherWebsite: (url: string) =>
              window.open(`${url}?utm_source=univer`, "_blank"),
          },
        }),
        UniverSheetsDataValidationPreset({
          showEditOnDropdown: true,
        }),
        UniverSheetsDrawingPreset(),
        UniverSheetsFindReplacePreset(),
      ],
    }) as CreateUniverResult;

    // Create workbook with data FIRST
    univerAPI.createWorkbook(initialData);
    univerAPIRef.current = univerAPI;
    univerRef.current = univer;

    const eventHandlers: { event: any; handler: any }[] = [];

    // Setup lifecycle event for images AFTER workbook is created
    univerAPI?.addEvent?.(
      univerAPI?.Event?.LifeCycleChanged,
      async (event: any) => {
        if (event.stage === univerAPI?.Enum?.LifecycleStages?.Rendered) {
          // Only add demo images if this is initial load (not import)
          if (initialData === WORKBOOK_DATA) {
            const fWorkbook = univerAPI?.getActiveWorkbook?.() as FWorkbook;
            const fWorksheet = (fWorkbook as FWorkbook)?.getActiveSheet();
            const imageUrl = "https://avatars.githubusercontent.com/u/61444807";

            try {
              const image = await fWorksheet
                ?.newOverGridImage()
                .setSource(imageUrl, univerAPI?.Enum?.ImageSourceType?.URL)
                .setColumn(5)
                .setRow(5)
                .setWidth(120)
                .setHeight(120)
                .buildAsync();

              if (image) {
                fWorksheet?.insertImages([image]);
              }

              // Insert cell images
              for (const cell of ["A11", "B12", "C13"]) {
                await fWorksheet
                  ?.getRange(cell)
                  ?.insertCellImageAsync(imageUrl);
              }
            } catch (error) {
              console.error("Error inserting demo images:", error);
            }
          }

          // Setup drag and drop after everything is rendered
          setTimeout(() => {
            setupDragAndDropEvents(univerAPI, eventHandlers);
          }, 100);
        }
      }
    );

    // Setup command monitoring
    const injector = univer.__getInjector();
    const commandService = injector.get(ICommandService);
    const extendedCommandService = commandService as ICommandService &
      ExtendedCommandService;

    const disposers: Array<() => void> = [];

    if (extendedCommandService.commandExecuted?.subscribe) {
      const subscription = extendedCommandService.commandExecuted.subscribe(
        throttle(() => {
          const active = univerAPIRef.current?.getActiveWorkbook?.();
          const wb: FWorkbook | undefined =
            active && typeof active !== "string"
              ? active
              : typeof active === "string"
              ? univerAPIRef.current?.getWorkbook?.(active)
              : undefined;

          if (wb && typeof wb.getSnapshot === "function") {
            console.log("[Workbook snapshot]", wb.getSnapshot());
          }
        }, 300)
      );
      disposers.push(() => subscription.unsubscribe());
    }

    return { univerAPI, univer, eventHandlers, disposers };
  };

  const setupDragAndDropEvents = (
    univerAPI: CreateUniverResult["univerAPI"],
    eventHandlers: { event: any; handler: any }[]
  ) => {
    let lastRow = -1;
    let lastCol = -1;

    const dragOverHandler = (params: ICellEventParam) => {
      const { worksheet, row, column } = params;
      (params.event as DragEvent | undefined)?.preventDefault();

      if (!worksheet) return;
      if (row === lastRow && column === lastCol) return;
      lastRow = row;
      lastCol = column;

      if (isDraggingRef.current) {
        const cellRef = worksheet.getRange(row, column).getA1Notation();
        setHoveredCell(cellRef);
      }
    };

    const dropHandler = async (params: ICellEventParam) => {
      const { worksheet, row, column } = params;
      const nativeEvent = params.event as DragEvent | undefined;

      nativeEvent?.preventDefault();

      if (!worksheet) return;

      const fromNative = collectDragData(nativeEvent?.dataTransfer);
      const fallback = draggedDataRef.current || {};
      const files =
        (fromNative.files.length ? fromNative.files : fallback.files) || [];
      const urls =
        (fromNative.urls.length ? fromNative.urls : fallback.urls) || [];

      if (files.length === 0 && urls.length === 0) {
        setIsDragging(false);
        setHoveredCell(null);
        draggedDataRef.current = null;
        return;
      }

      // Insert images sequentially
      let insertedCount = 0;

      for (let i = 0; i < files.length; i++) {
        const targetRow = row + i;
        const a1 = worksheet.getRange(targetRow, column).getA1Notation();
        await insertImageToCell(files[i], a1);
        insertedCount++;
      }

      const startOffset = files.length;
      for (let i = 0; i < urls.length; i++) {
        const targetRow = row + startOffset + i;
        const a1 = worksheet.getRange(targetRow, column).getA1Notation();
        await insertImageUrlToCell(urls[i], a1);
        insertedCount++;
      }

      draggedDataRef.current = null;
      setIsDragging(false);
      setHoveredCell(null);

      if (insertedCount > 0) {
        showNotification(`Successfully inserted ${insertedCount} image(s)`);
      }
    };

    if (univerAPI?.Event?.DragOver && univerAPI?.addEvent) {
      univerAPI.addEvent(univerAPI.Event.DragOver, dragOverHandler);
      eventHandlers.push({
        event: univerAPI.Event.DragOver,
        handler: dragOverHandler,
      });
    }

    if (univerAPI?.Event?.Drop && univerAPI?.addEvent) {
      univerAPI.addEvent(univerAPI.Event.Drop, dropHandler);
      eventHandlers.push({
        event: univerAPI.Event.Drop,
        handler: dropHandler,
      });
    }
  };

  const setupNativeDragHandlers = () => {
    const container = containerRef.current;
    if (!container) return;

    const handleDragOver = (e: DragEvent) => {
      e.preventDefault();
      if (e.dataTransfer) e.dataTransfer.dropEffect = "copy";
      setDragPos({ x: e.clientX, y: e.clientY });
    };

    const handleDragEnter = (e: DragEvent) => {
      e.preventDefault();
      setIsDragging(true);
    };

    const handleDragLeave = (e: DragEvent) => {
      e.preventDefault();
      const rect = container.getBoundingClientRect();
      const x = e.clientX;
      const y = e.clientY;
      if (
        x <= rect.left ||
        x >= rect.right ||
        y <= rect.top ||
        y >= rect.bottom
      ) {
        setIsDragging(false);
        setHoveredCell(null);
        setDragPos(null);
      }
    };

    const handleDrop = (e: DragEvent) => {
      e.preventDefault();
      setDragPos({ x: e.clientX, y: e.clientY });

      const data = collectDragData(e.dataTransfer);
      draggedDataRef.current = {
        files: data.files,
        urls: data.urls,
      };
    };

    container.addEventListener("dragover", handleDragOver, { passive: false });
    container.addEventListener("dragenter", handleDragEnter, {
      passive: false,
    });
    container.addEventListener("dragleave", handleDragLeave, {
      passive: false,
    });
    container.addEventListener("drop", handleDrop, {
      passive: false,
      capture: true,
    });

    return () => {
      container.removeEventListener("dragover", handleDragOver);
      container.removeEventListener("dragenter", handleDragEnter);
      container.removeEventListener("dragleave", handleDragLeave);
      container.removeEventListener("drop", handleDrop, true);
    };
  };

  useEffect(() => {
    // Initialize Univer with default data
    const result = initializeUniver();

    // Setup native drag handlers
    const cleanupNativeHandlers = setupNativeDragHandlers();

    // Cleanup function
    return () => {
      // Remove event handlers
      if (result.eventHandlers) {
        result.eventHandlers.forEach(({ event, handler }) => {
          if (univerAPIRef.current?.removeEvent) {
            univerAPIRef.current.removeEvent(event, handler);
          }
        });
      }

      // Cleanup native handlers
      if (cleanupNativeHandlers) {
        cleanupNativeHandlers();
      }

      // Dispose subscriptions
      if (result.disposers) {
        result.disposers.forEach((dispose) => dispose());
      }

      // Dispose Univer instance
      if (univerAPIRef.current?.dispose) {
        univerAPIRef.current.dispose();
      }

      univerAPIRef.current = null;
      univerRef.current = null;
    };
  }, [isDark]);

  return (
    <div className="w-screen h-screen dark">
      {/* Export/Import Buttons */}
      <div className="fixed top-0.5 right-4 z-[9999] flex gap-2">
        <Button
          onClick={handleExport}
          variant="secondary"
          size="sm"
          className="shadow-md"
        >
          📥 Export
        </Button>
        <Button
          onClick={handleImport}
          variant="secondary"
          size="sm"
          className="shadow-md"
        >
          📤 Import
        </Button>
      </div>

      {/* Notification Toast */}
      {notification && (
        <div className="fixed top-[70px] right-4 z-[10000] px-5 py-3 bg-green-600 text-white rounded shadow-lg animate-in slide-in-from-right-full duration-300">
          {notification}
        </div>
      )}

      {/* Spreadsheet Container */}
      <div
        ref={containerRef}
        className={cn(
          "w-full h-full",
          isDragging &&
            "outline-2 outline-dashed outline-blue-500 outline-offset-[-2px]"
        )}
      />

      {/* Floating drag feedback near cursor */}
      {isDragging && dragPos && (
        <div
          className="fixed pointer-events-none z-[10000] px-3 py-2 bg-background text-white rounded text-xs shadow-lg select-none"
          style={{
            left: dragPos.x + 30,
            top: dragPos.y + 20,
            transform: "translate(8px, 8px)",
          }}
        >
          {hoveredCell
            ? `Drop to insert image in ${hoveredCell}`
            : "Drag over a cell…"}
        </div>
      )}
    </div>
  );
}

export default App;
