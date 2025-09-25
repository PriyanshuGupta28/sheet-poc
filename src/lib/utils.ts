import { clsx, type ClassValue } from "clsx";
import { twMerge } from "tailwind-merge";

export function cn(...inputs: ClassValue[]) {
  return twMerge(clsx(inputs));
}

// Throttle helper with strict typing
export function throttle<Args extends unknown[]>(
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
