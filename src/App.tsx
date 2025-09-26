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
    if (containerRef.current) {
      univerRef.current = univer; // Save Univer instance
    }

    // Get CommandService
    const injector = univer.__getInjector();
    const commandService = injector.get(ICommandService);
    const extendedCommandService = commandService as ICommandService &
      ExtendedCommandService;

    // Subscribe AFTER command execution
    const disposers: Array<() => void> = [];
    const onAfter = (handler: (event: CommandExecutedEvent) => void) => {
      if (extendedCommandService.commandExecuted$?.subscribe) {
        const subscription: Subscription =
          extendedCommandService.commandExecuted$.subscribe(handler);
        disposers.push(() => subscription.unsubscribe());
      }
    };

    // Throttled full-state dump
    const dumpStateThrottled = throttle(() => {
      const active = univerAPI.getActiveWorkbook?.();

      const wb: FWorkbook | undefined =
        active && typeof active !== "string"
          ? active
          : typeof active === "string"
          ? univerAPI.getWorkbook?.(active)
          : undefined;

      if (!wb) return;

      // Try different methods to get snapshot
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
        // Pretty-print once per burst of changes
        console.log("[Workbook snapshot]", snapshot);
        // Or: console.log(JSON.stringify(snapshot, null, 2));
      } else {
        console.warn("Could not obtain workbook snapshot; inspect wb:", wb);
      }
    }, 300);

    onAfter(() => {
      dumpStateThrottled();
    });

    return () => {
      // Dispose of all subscriptions
      disposers.forEach((dispose) => dispose());
      univerAPI.dispose();
    };
  }, [isDark]);

  return (
    <div style={{ width: "100vw", height: "100vh" }}>
      <div ref={containerRef} style={{ width: "100%", height: "100%" }} />
    </div>
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
