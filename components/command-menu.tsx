"use client";

import { useEffect, useRef, type ReactNode } from "react";

export function CommandMenu({ children, label, onOpenChange, open }: {
  children: ReactNode;
  label: string;
  onOpenChange: (open: boolean) => void;
  open: boolean;
}) {
  const rootRef = useRef<HTMLDetailsElement>(null);
  const summaryRef = useRef<HTMLElement>(null);

  useEffect(() => {
    if (!open) return;
    function handlePointerDown(event: PointerEvent) {
      if (!rootRef.current?.contains(event.target as Node)) onOpenChange(false);
    }
    function handleKeyDown(event: KeyboardEvent) {
      if (event.key !== "Escape") return;
      event.preventDefault();
      onOpenChange(false);
      summaryRef.current?.focus();
    }
    document.addEventListener("pointerdown", handlePointerDown);
    document.addEventListener("keydown", handleKeyDown);
    return () => {
      document.removeEventListener("pointerdown", handlePointerDown);
      document.removeEventListener("keydown", handleKeyDown);
    };
  }, [onOpenChange, open]);

  return (
    <details
      ref={rootRef}
      className="commandMenu"
      open={open}
      onToggle={(event) => onOpenChange(event.currentTarget.open)}
      onClick={(event) => {
        if ((event.target as HTMLElement).closest("button")) onOpenChange(false);
      }}
    >
      <summary ref={summaryRef}>{label}</summary>
      <div>{children}</div>
    </details>
  );
}
