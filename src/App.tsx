/* eslint-disable @typescript-eslint/no-unused-expressions */
import "./App.css";
import { useEffect, useRef, useState } from "react";
import {
  createUniver,
  ICommandService,
  LocaleType,
  mergeLocales,
  Univer,
} from "@univerjs/presets";

// presets
import { UniverSheetsHyperLinkPreset } from "@univerjs/preset-sheets-hyper-link";
import { UniverSheetsCorePreset } from "@univerjs/preset-sheets-core";
import { UniverSheetsDataValidationPreset } from "@univerjs/preset-sheets-data-validation";
import { UniverSheetsDrawingPreset } from "@univerjs/preset-sheets-drawing";

// locales
import UniverPresetSheetsCoreEnUS from "@univerjs/preset-sheets-core/locales/en-US";
import UniverPresetSheetsHyperLinkEnUS from "@univerjs/preset-sheets-hyper-link/locales/en-US";
import UniverPresetSheetsDataValidationEnUS from "@univerjs/preset-sheets-data-validation/locales/en-US";
import UniverPresetSheetsDrawingEnUS from "@univerjs/preset-sheets-drawing/locales/en-US";

// css
import "@univerjs/preset-sheets-core/lib/index.css";
import "@univerjs/preset-sheets-hyper-link/lib/index.css";
import "@univerjs/preset-sheets-data-validation/lib/index.css";
import "@univerjs/preset-sheets-drawing/lib/index.css";

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
    getSheetHooks?: () => any;
  };
  univer: Univer;
}

interface CommandExecutedEvent {
  id: string;
  type?: string;
  params?: Record<string, unknown>;
}

interface ExtendedCommandService {
  commandExecuted$?: {
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

  // Use arrays for files to avoid FileList quirks
  const draggedDataRef = useRef<{ files?: File[]; urls?: string[] } | null>(
    null
  );

  // Highlight state
  const currentHoverRef = useRef<{ row: number; col: number } | null>(null);
  const insertedHighlightsRef = useRef<
    Map<string, { row: number; col: number; until: number }>
  >(new Map());
  const renderHandlerRef = useRef<any>(null);
  const rafRefreshRef = useRef<number | null>(null);
  const rafTickRef = useRef<number | null>(null);

  const isDark =
    typeof document !== "undefined" &&
    document.documentElement.classList.contains("dark");

  // Prevent browser navigation on file drop outside the grid
  useEffect(() => {
    const preventNav = (e: DragEvent) => {
      e.preventDefault();
      // Do NOT stopPropagation here; let events flow to elements below
    };
    window.addEventListener("dragover", preventNav, { passive: false });
    window.addEventListener("drop", preventNav, { passive: false });
    return () => {
      window.removeEventListener("dragover", preventNav);
      window.removeEventListener("drop", preventNav);
    };
  }, []);

  useEffect(() => {
    const { univerAPI, univer } = createUniver({
      darkMode: isDark,
      locale: LocaleType.EN_US,
      locales: {
        [LocaleType.EN_US]: mergeLocales(
          UniverPresetSheetsCoreEnUS,
          UniverPresetSheetsHyperLinkEnUS,
          UniverPresetSheetsDataValidationEnUS,
          UniverPresetSheetsDrawingEnUS
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
      ],
    }) as CreateUniverResult;

    univerAPI.createWorkbook(WORKBOOK_DATA);
    univerAPIRef.current = univerAPI;

    const eventHandlers: { event: any; handler: any }[] = [];

    univerAPI?.addEvent?.(
      univerAPI?.Event?.LifeCycleChanged,
      async (event: any) => {
        if (event.stage === univerAPI?.Enum?.LifecycleStages?.Rendered) {
          // Example starter images (safe to keep or remove)
          const fWorkbook = univerAPI?.getActiveWorkbook?.();
          const fWorksheet = (fWorkbook as any)?.getActiveSheet();
          const imageUrl = "https://avatars.githubusercontent.com/u/61444807";

          const image = await fWorksheet
            ?.newOverGridImage()
            .setSource(imageUrl, univerAPI?.Enum?.ImageSourceType?.URL)
            .setColumn(5)
            .setRow(5)
            .setWidth(120)
            .setHeight(120)
            .buildAsync();
          image && fWorksheet?.insertImages([image]);

          ["A11", "B12", "C13"].forEach((cell) =>
            fWorksheet?.getRange(cell)?.insertCellImageAsync(imageUrl)
          );

          setTimeout(() => {
            setupDragAndDropEvents();
          }, 0);
        }
      }
    );

    if (containerRef.current) {
      univerRef.current = univer;
    }

    // Rendering helpers
    const scheduleCanvasRefresh = () => {
      if (rafRefreshRef.current != null) return;
      rafRefreshRef.current = requestAnimationFrame(() => {
        rafRefreshRef.current = null;
        const workbook = univerAPI?.getActiveWorkbook?.();
        if (workbook && typeof workbook !== "string") {
          const worksheet = (workbook as any).getActiveSheet?.();
          if (worksheet && typeof worksheet.refreshCanvas === "function") {
            worksheet.refreshCanvas();
          }
        }
      });
    };

    const ensureRenderHandler = () => {
      const hooks = univerAPI?.getSheetHooks?.();
      if (!hooks || renderHandlerRef.current) return;

      renderHandlerRef.current = {
        drawWith: (ctx: CanvasRenderingContext2D, info: any) => {
          const hover = currentHoverRef.current;
          const key = `${info.row},${info.col}`;
          const now = performance.now();
          const inserted = insertedHighlightsRef.current.get(key);
          const { primaryWithCoord } = info;
          if (!primaryWithCoord) return;

          const { startX, startY, endX, endY } = primaryWithCoord;
          const padding = 2;

          // Inserted highlight (green, solid, ✓ pulse)
          if (inserted && inserted.until > now) {
            ctx.save();
            ctx.strokeStyle = "#28a745";
            ctx.fillStyle = "rgba(40, 167, 69, 0.18)";
            ctx.lineWidth = 3;
            ctx.setLineDash([]);
            ctx.fillRect(
              startX - padding,
              startY - padding,
              endX - startX + padding * 2,
              endY - startY + padding * 2
            );
            ctx.strokeRect(
              startX - padding,
              startY - padding,
              endX - startX + padding * 2,
              endY - startY + padding * 2
            );
            ctx.fillStyle = "#28a745";
            ctx.font = "bold 14px sans-serif";
            ctx.fillText("✓", startX + 6, startY + 16);
            ctx.restore();
          }

          // Hover highlight (blue, dashed)
          if (hover && info.row === hover.row && info.col === hover.col) {
            ctx.save();
            ctx.strokeStyle = "#007bff";
            ctx.fillStyle = "rgba(0, 123, 255, 0.12)";
            ctx.lineWidth = 3;
            ctx.setLineDash([5, 5]);
            ctx.fillRect(
              startX - padding,
              startY - padding,
              endX - startX + padding * 2,
              endY - startY + padding * 2
            );
            ctx.strokeRect(
              startX - padding,
              startY - padding,
              endX - startX + padding * 2,
              endY - startY + padding * 2
            );
            ctx.restore();
          }
        },
      };

      hooks.onCellRender([renderHandlerRef.current]);
      scheduleCanvasRefresh();
    };

    const startTick = () => {
      if (rafTickRef.current != null) return;
      const tick = () => {
        const now = performance.now();
        let changed = false;
        for (const [k, v] of insertedHighlightsRef.current) {
          if (v.until <= now) {
            insertedHighlightsRef.current.delete(k);
            changed = true;
          }
        }
        if (changed) scheduleCanvasRefresh();
        if (insertedHighlightsRef.current.size > 0) {
          rafTickRef.current = requestAnimationFrame(tick);
        } else {
          rafTickRef.current = null;
        }
      };
      rafTickRef.current = requestAnimationFrame(tick);
    };

    const highlightHoverCell = (row: number, col: number) => {
      ensureRenderHandler();
      const prev = currentHoverRef.current;
      if (!prev || prev.row !== row || prev.col !== col) {
        currentHoverRef.current = { row, col };
        scheduleCanvasRefresh();
      }
    };

    const clearHoverHighlight = () => {
      currentHoverRef.current = null;
      scheduleCanvasRefresh();
    };

    const addInsertedHighlight = (row: number, col: number, ms = 1200) => {
      ensureRenderHandler();
      const key = `${row},${col}`;
      insertedHighlightsRef.current.set(key, {
        row,
        col,
        until: performance.now() + ms,
      });
      startTick();
      scheduleCanvasRefresh();
    };

    // Data extraction helpers
    const extractImageUrls = (html: string): string[] => {
      const parser = new DOMParser();
      const doc = parser.parseFromString(html, "text/html");
      return Array.from(doc.querySelectorAll("img"))
        .map((img) => img.src)
        .filter(
          (src) =>
            src && (src.startsWith("http") || src.startsWith("data:image"))
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
      } catch {}

      try {
        const html = dt.getData("text/html");
        if (html) out.urls.push(...extractImageUrls(html));
      } catch {}

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
      } catch {}

      try {
        const plain = dt.getData("text/plain");
        if (
          plain &&
          (plain.startsWith("http") || plain.startsWith("data:image"))
        ) {
          out.urls.push(plain.trim());
        }
      } catch {}

      // Dedup urls
      out.urls = Array.from(new Set(out.urls));
      return out;
    };

    // Univer drag & drop events (use captured data if native event dt is missing)
    const setupDragAndDropEvents = () => {
      let lastRow = -1;
      let lastCol = -1;

      const dragOverHandler = (params: ICellEventParam) => {
        const { worksheet, row, column } = params;
        // Always preventDefault to allow drop
        (params.event as DragEvent | undefined)?.preventDefault();

        if (!worksheet) return;
        if (row === lastRow && column === lastCol) return;
        lastRow = row;
        lastCol = column;

        if (isDraggingRef.current) {
          const cellRef = worksheet.getRange(row, column).getA1Notation();
          setHoveredCell(cellRef);
          highlightHoverCell(row, column);
        }
      };

      const dropHandler = async (params: ICellEventParam) => {
        const { worksheet, row, column } = params;
        const nativeEvent = params.event as DragEvent | undefined;

        nativeEvent?.preventDefault();

        if (!worksheet) return;

        // Prefer native dt; otherwise fall back to container-captured data
        const fromNative = collectDragData(nativeEvent?.dataTransfer);
        const fallback = draggedDataRef.current || {};
        const files =
          (fromNative.files.length ? fromNative.files : fallback.files) || [];
        const urls =
          (fromNative.urls.length ? fromNative.urls : fallback.urls) || [];

        if (files.length === 0 && urls.length === 0) {
          // No data; just clear UI
          setIsDragging(false);
          setHoveredCell(null);
          clearHoverHighlight();
          draggedDataRef.current = null;
          return;
        }

        // Insert files row-by-row
        for (let i = 0; i < files.length; i++) {
          const targetRow = row + i;
          const a1 = worksheet.getRange(targetRow, column).getA1Notation();
          await insertImageToCell(files[i], a1);
          addInsertedHighlight(targetRow, column);
        }

        // Insert URLs after files
        const startOffset = files.length;
        for (let i = 0; i < urls.length; i++) {
          const targetRow = row + startOffset + i;
          const a1 = worksheet.getRange(targetRow, column).getA1Notation();
          await insertImageUrlToCell(urls[i], a1);
          addInsertedHighlight(targetRow, column);
        }

        // Cleanup UI state
        draggedDataRef.current = null;
        setIsDragging(false);
        setHoveredCell(null);
        clearHoverHighlight();
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

    // Native drag handlers — UI only + capture data before Univer Drop
    const setupNativeDragHandlers = () => {
      const container = containerRef.current;
      if (!container) return;

      const handleDragOver = (e: DragEvent) => {
        e.preventDefault(); // allow drop
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
          clearHoverHighlight();
          setDragPos(null);
        }
      };

      // IMPORTANT: capture: true so this runs BEFORE Univer's drop handler
      const handleDrop = (e: DragEvent) => {
        e.preventDefault(); // do not stopPropagation — we want Univer to receive it
        setDragPos({ x: e.clientX, y: e.clientY });

        const data = collectDragData(e.dataTransfer);
        draggedDataRef.current = {
          files: data.files,
          urls: data.urls,
        };
      };

      container.addEventListener("dragover", handleDragOver, {
        passive: false,
      });
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

    // Insert image file to cell
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

    // Insert image from URL to cell
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
        }
      } catch (error) {
        console.error("Error inserting image from URL:", error);
      }
    };

    // CommandService debug (unchanged)
    const injector = univer.__getInjector();
    const commandService = injector.get(ICommandService);
    const extendedCommandService = commandService as ICommandService &
      ExtendedCommandService;

    const disposers: Array<() => void> = [];
    const onAfter = (handler: (event: CommandExecutedEvent) => void) => {
      if (extendedCommandService.commandExecuted$?.subscribe) {
        const subscription: Subscription =
          extendedCommandService.commandExecuted$.subscribe(handler);
        disposers.push(() => subscription.unsubscribe());
      }
    };

    const dumpStateThrottled = throttle(() => {
      const active = univerAPIRef.current?.getActiveWorkbook?.();
      const wb: FWorkbook | undefined =
        active && typeof active !== "string"
          ? active
          : typeof active === "string"
          ? univerAPIRef.current?.getWorkbook?.(active)
          : undefined;
      if (!wb) return;

      let snapshot: unknown;
      if (typeof wb.getSnapshot === "function") snapshot = wb.getSnapshot();
      else if (
        typeof (wb as FWorkbook & { serialize: () => unknown }).serialize ===
        "function"
      )
        snapshot = (wb as FWorkbook & { serialize: () => unknown }).serialize();
      else if (typeof wb.save === "function") snapshot = wb.save();
      else if (
        typeof (wb as FWorkbook & { toJSON: () => unknown }).toJSON ===
        "function"
      )
        snapshot = (wb as FWorkbook & { toJSON: () => unknown }).toJSON();
      else if (
        typeof (wb as FWorkbook & { toJson: () => unknown }).toJson ===
        "function"
      )
        snapshot = (wb as FWorkbook & { toJson: () => unknown }).toJson();

      if (snapshot) console.log("[Workbook snapshot]", snapshot);
    }, 300);

    onAfter(() => {
      dumpStateThrottled();
    });

    const cleanupNativeHandlers = setupNativeDragHandlers();

    // Cleanup
    return () => {
      eventHandlers.forEach(({ event, handler }) => {
        if (univerAPIRef.current?.removeEvent) {
          univerAPIRef.current.removeEvent(event, handler);
        }
      });
      if (cleanupNativeHandlers) cleanupNativeHandlers();
      try {
        const hooks = univerAPIRef.current?.getSheetHooks?.();
        if (hooks) hooks.onCellRender([]);
      } catch {}
      if (rafRefreshRef.current != null)
        cancelAnimationFrame(rafRefreshRef.current);
      if (rafTickRef.current != null) cancelAnimationFrame(rafTickRef.current);
      if (univerAPIRef.current?.dispose) {
        univerAPIRef.current.dispose();
      }
    };
  }, [isDark]); // do NOT depend on isDragging

  return (
    <>
      <style>
        {`
        .drag-feedback {
          position: fixed;
          pointer-events: none;
          z-index: 10000;
          padding: 8px 12px;
          background: rgba(0, 123, 255, 0.9);
          color: white;
          border-radius: 4px;
          font-size: 12px;
          box-shadow: 0 2px 8px rgba(0, 0, 0, 0.2);
          transform: translate(8px, 8px);
          user-select: none;
        }

        .container--dragging {
          outline: 2px dashed #007bff;
          outline-offset: -2px;
        }

        canvas { position: relative; z-index: 1; }
      `}
      </style>

      <div style={{ width: "100vw", height: "100vh" }} className="dark">
        <div
          ref={containerRef}
          style={{ width: "100%", height: "100%" }}
          className={isDragging ? "container--dragging" : ""}
        />

        {/* Instructions overlay */}
        <div
          style={{
            position: "fixed",
            top: 10,
            right: 10,
            background: "rgba(0, 0, 0, 0.8)",
            color: "white",
            padding: "10px 15px",
            borderRadius: "5px",
            fontSize: "14px",
            zIndex: 1000,
          }}
        >
          <div>Drag and drop images directly onto any cell</div>
          {hoveredCell && isDragging && (
            <div style={{ marginTop: "5px", color: "#4CAF50" }}>
              Drop target: {hoveredCell}
            </div>
          )}
          {isDragging && !hoveredCell && (
            <div style={{ marginTop: "5px", color: "#FFC107" }}>
              Drag over a cell to select it
            </div>
          )}
        </div>

        {/* Floating drag feedback near cursor */}
        {isDragging && dragPos && (
          <div
            className="drag-feedback"
            style={{ left: dragPos.x, top: dragPos.y }}
          >
            {hoveredCell
              ? `Drop to insert image in ${hoveredCell}`
              : "Drag over a cell…"}
          </div>
        )}
      </div>
    </>
  );
}

export default App;
