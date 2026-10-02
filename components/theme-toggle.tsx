"use client";

import { Moon, Sun } from "lucide-react";
import { Button } from "@/components/ui/button";

export function ThemeToggle() {
  return (
    <Button
      type="button"
      variant="outline"
      size="icon"
      aria-label="Сменить тему"
      onClick={() => {
        const dark = !document.documentElement.classList.contains("dark");
        document.documentElement.classList.toggle("dark", dark);
        localStorage.setItem("sverka.theme", dark ? "dark" : "light");
      }}
    >
      <Sun className="hidden dark:block" />
      <Moon className="dark:hidden" />
    </Button>
  );
}
