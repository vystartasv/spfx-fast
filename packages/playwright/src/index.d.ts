import type { Page } from "playwright";

export type RecipeLocator =
  | { role: string; name: string; exact?: boolean }
  | { text: string; exact?: boolean }
  | { label: string; exact?: boolean }
  | { testId: string }
  | { placeholder: string; exact?: boolean };

export type RecipeStep =
  | { action: "navigate"; path: string; query?: Record<string, string> }
  | { action: "assertPageReady" }
  | { action: "assertAuthenticated" }
  | { action: "click"; locator: RecipeLocator; confirm?: true }
  | { action: "assertVisible"; locator: RecipeLocator }
  | { action: "fill"; locator: RecipeLocator; value: string }
  | { action: "select"; locator: RecipeLocator; value: string }
  | { action: "setChecked"; locator: RecipeLocator; checked: boolean }
  | { action: "assertText" | "assertValue"; locator: RecipeLocator; expected: string; exact?: boolean }
  | { action: "screenshot"; name: string; fullPage?: boolean };

export interface SharePointRecipe {
  name: string;
  steps: RecipeStep[];
}

export function validateRecipe(recipe: unknown): SharePointRecipe;
export function validateBaseUrl(baseUrl: string): URL;
export function runRecipe(page: Page, recipe: SharePointRecipe, options: {
  baseUrl: string;
  screenshotDir?: string;
  timeoutMs?: number;
  fast?: boolean;
}): Promise<{ name: string; steps: number; screenshots: string[] }>;
export function generateTestSource(recipe: SharePointRecipe): string;
