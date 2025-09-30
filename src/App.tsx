/* eslint-disable @typescript-eslint/no-unused-expressions */
import "./App.css";
import { useEffect, useRef } from "react";
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

interface CreateUniverResult {
  univerAPI: {
    createWorkbook: (data: unknown) => void;
    getActiveWorkbook?: () => FWorkbook | string | undefined;
    getWorkbook?: (id: string) => FWorkbook | undefined;
    dispose: () => void;
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

function App() {
  const containerRef = useRef<HTMLDivElement>(null);
  const univerRef = useRef<Univer | null>(null);
  const univerAPIRef = useRef<any>(null);

  const isDark =
    typeof document !== "undefined" &&
    document.documentElement.classList.contains("dark");

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

    // Your existing lifecycle event handler
    univerAPI?.addEvent(univerAPI?.Event.LifeCycleChanged, async (event) => {
      if (event.stage === univerAPI.Enum.LifecycleStages.Rendered) {
        const fWorkbook = univerAPI?.getActiveWorkbook();
        const fWorksheet = fWorkbook?.getActiveSheet();

        const imageUrl = "https://avatars.githubusercontent.com/u/61444807";

        // Insert a floating image into the active worksheet
        const image = await fWorksheet
          ?.newOverGridImage()
          .setSource(imageUrl, univerAPI?.Enum.ImageSourceType.URL)
          .setColumn(5)
          .setRow(5)
          .setWidth(120)
          .setHeight(120)
          .buildAsync();

        image && fWorksheet?.insertImages([image]);

        // Insert a cell image into the active worksheet
        const cells = [
          "A11",
          "B12",
          "C13",
          "D14",
          "E15",
          "F16",
          "G17",
          "H18",
          "I19",
          "J20",
        ];
        cells.forEach((cell) => {
          const fRange = fWorksheet?.getRange(cell);
          fRange?.insertCellImageAsync(imageUrl);
        });

        // Setup drag and drop after render
        setupImageDragAndDrop();
      }
    });

    // Image drag and drop setup
    const setupImageDragAndDrop = () => {
      const container = containerRef.current;
      if (!container) return;

      let dragCounter = 0; // To handle drag enter/leave properly

      // Prevent default drag behaviors
      container.addEventListener("dragenter", (e) => {
        e.preventDefault();
        dragCounter++;
        container.classList.add("drag-over");
      });

      container.addEventListener("dragover", (e) => {
        e.preventDefault();
        e.stopPropagation();

        // Show copy cursor
        if (e.dataTransfer) {
          e.dataTransfer.dropEffect = "copy";
        }
      });

      container.addEventListener("dragleave", (e) => {
        e.preventDefault();
        dragCounter--;
        if (dragCounter === 0) {
          container.classList.remove("drag-over");
        }
      });

      container.addEventListener("drop", async (e) => {
        e.preventDefault();
        e.stopPropagation();

        dragCounter = 0;
        container.classList.remove("drag-over");

        const dt = e.dataTransfer;
        if (!dt) return;

        // Handle dropped files
        if (dt.files && dt.files.length > 0) {
          const imageFiles = Array.from(dt.files).filter((file) =>
            file.type.startsWith("image/")
          );

          if (imageFiles.length === 0) {
            console.log("No image files detected in drop");
            return;
          }

          // Get the target cell based on drop position
          const cellInfo = await getCellFromDropPosition(e, container);

          if (!cellInfo) {
            console.log("Could not determine target cell");
            return;
          }

          console.log(
            `Dropping image to cell: ${cellInfo.cellRef} (Row: ${cellInfo.row}, Col: ${cellInfo.col})`
          );

          // Process each image
          for (const file of imageFiles) {
            await insertImageToCell(file, cellInfo.cellRef);
          }
        }

        // Handle dragged images from web pages
        const html = dt.getData("text/html");
        if (html) {
          const imgUrls = extractImageUrls(html);
          if (imgUrls.length > 0) {
            const cellInfo = await getCellFromDropPosition(e, container);
            if (cellInfo) {
              for (const url of imgUrls) {
                await insertImageUrlToCell(url, cellInfo.cellRef);
              }
            }
          }
        }
      });
    };

    // Get cell information from drop position
    const getCellFromDropPosition = async (
      e: DragEvent,
      container: HTMLElement
    ) => {
      const workbook = univerAPI.getActiveWorkbook?.();
      if (!workbook || typeof workbook === "string") return null;

      const sheet = workbook.getActiveSheet();
      if (!sheet) return null;

      // Method 1: Try to use the selected cell
      const selection = sheet.getSelection();
      if (selection) {
        const range = selection.getActiveRange();
        const row = range?.getRow();
        const col = range?.getColumn();
        if (row === undefined || col === undefined) return null;
        const cellRef = `${String.fromCharCode(65 + col)}${row + 1}`;

        return { row, col, cellRef };
      }

      // Method 2: Calculate from mouse position (simplified)
      const rect = container.getBoundingClientRect();
      const x = e.clientX - rect.left;
      const y = e.clientY - rect.top;

      // These are approximate values - adjust based on your UI
      const HEADER_HEIGHT = 30;
      const ROW_HEIGHT = 25;
      const COL_WIDTH = 100;
      const ROW_HEADER_WIDTH = 50;

      if (y < HEADER_HEIGHT || x < ROW_HEADER_WIDTH) {
        return null; // Dropped on headers
      }

      const col = Math.floor((x - ROW_HEADER_WIDTH) / COL_WIDTH);
      const row = Math.floor((y - HEADER_HEIGHT) / ROW_HEIGHT);

      // Ensure valid cell coordinates
      if (col < 0 || row < 0 || col > 25 || row > 999) {
        return null;
      }

      const cellRef = `${String.fromCharCode(65 + col)}${row + 1}`;

      return { row, col, cellRef };
    };

    // Insert image file to cell
    const insertImageToCell = async (file: File, cellRef: string) => {
      return new Promise<void>((resolve) => {
        const reader = new FileReader();

        reader.onload = async (e) => {
          const dataUrl = e.target?.result as string;

          try {
            const workbook = univerAPI.getActiveWorkbook?.();
            if (!workbook || typeof workbook === "string") {
              resolve();
              return;
            }

            const sheet = workbook.getActiveSheet();
            if (!sheet) {
              resolve();
              return;
            }

            const range = sheet.getRange(cellRef);
            if (range) {
              // Insert as cell image
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
        const workbook = univerAPI.getActiveWorkbook?.();
        if (!workbook || typeof workbook === "string") return;

        const sheet = workbook.getActiveSheet();
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

    // Extract image URLs from HTML
    const extractImageUrls = (html: string): string[] => {
      const parser = new DOMParser();
      const doc = parser.parseFromString(html, "text/html");
      const images = doc.querySelectorAll("img");
      return Array.from(images)
        .map((img) => img.src)
        .filter(
          (src) => src && (src.startsWith("http") || src.startsWith("data:"))
        );
    };

    // Your existing command service setup
    if (containerRef.current) {
      univerRef.current = univer;
    }

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
      const active = univerAPI.getActiveWorkbook?.();

      const wb: FWorkbook | undefined =
        active && typeof active !== "string"
          ? active
          : typeof active === "string"
          ? univerAPI.getWorkbook?.(active)
          : undefined;

      if (!wb) return;

      let snapshot: unknown;

      if (typeof wb.getSnapshot === "function") {
        snapshot = wb.getSnapshot();
      } else if (
        typeof (wb as FWorkbook & { serialize: () => unknown }).serialize ===
        "function"
      ) {
        snapshot = (wb as FWorkbook & { serialize: () => unknown }).serialize();
      } else if (typeof wb.save === "function") {
        snapshot = wb.save();
      } else if (
        typeof (wb as FWorkbook & { toJSON: () => unknown }).toJSON ===
        "function"
      ) {
        snapshot = (wb as FWorkbook & { toJSON: () => unknown }).toJSON();
      } else if (
        typeof (wb as FWorkbook & { toJson: () => unknown }).toJson ===
        "function"
      ) {
        snapshot = (wb as FWorkbook & { toJson: () => unknown }).toJson();
      }

      if (snapshot) {
        console.log("[Workbook snapshot]", snapshot);
      } else {
        console.warn("Could not obtain workbook snapshot; inspect wb:", wb);
      }
    }, 300);

    onAfter(() => {
      dumpStateThrottled();
    });

    return () => {
      disposers.forEach((dispose) => dispose());
      univerAPI.dispose();
    };
  }, [isDark]);

  return (
    <>
      <style>
        {`
          .drag-over {
            position: relative;
          }
          .drag-over::after {
            content: '';
            position: absolute;
            top: 0;
            left: 0;
            right: 0;
            bottom: 0;
            background-color: rgba(0, 123, 255, 0.1);
            border: 2px dashed #007bff;
            pointer-events: none;
            z-index: 1000;
          }
        `}
      </style>
      <div style={{ width: "100vw", height: "100vh" }} className="dark">
        <div ref={containerRef} style={{ width: "100%", height: "100%" }} />{" "}
        {/* Instructions overlay
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
          Drag and drop images from your computer or web browser onto cells
        </div> */}
      </div>
    </>
  );
}

export default App;

// Throttle helper with strict typing
function throttle<Args extends unknown[]>(
  fn: (...args: Args) => void,
  wait = 200
): (...args: Args) => void {
  let timeout: ReturnType<typeof setTimeout> | null = null;
  let pendingArgs: Args | null = null;

  return (...args: Args) => {
    if (timeout) {
      pendingArgs = args;
      return;
    }
    fn(...args);
    timeout = setTimeout(() => {
      timeout = null;
      if (pendingArgs) {
        const toRun = pendingArgs;
        pendingArgs = null;
        fn(...toRun);
      }
    }, wait);
  };
}
