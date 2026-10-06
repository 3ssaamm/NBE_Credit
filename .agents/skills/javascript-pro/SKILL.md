---
name: javascript-pro
description: >-
  Provides advanced standards and architectural patterns for modern JavaScript and TypeScript development, particularly in Node.js and Google Apps Script V8 runtime environments. Use when authoring, refactoring, or reviewing JavaScript/TypeScript code to enforce modern ES6+ standards, clean async/promise handling, modular structure, type safety with JSDoc/TypeScript, and defensive coding without outdated ES5 idioms or spaghetti antipatterns.
---

# JavaScript & TypeScript Professional Engineering Standards

This skill guides the design, implementation, refactoring, and code review of JavaScript and TypeScript codebases. It ensures clean, maintainable, modern, and high-performance code across both standard runtime environments and Google Apps Script (V8 runtime).

---

## 1. Core Language Standards (ES6+)

### Adopt Modern ECMAScript Features
- **Scoping**: Use `const` by default. Use `let` only when variable reassignment is strictly required. Never use `var`.
- **Destructuring**: Use object and array destructuring for clarity:
  ```javascript
  const { statementDate, billedBalance } = cardBalance;
  const [firstItem, ...restItems] = itemsList;
  ```
- **Optional Chaining & Nullish Coalescing**:
  - Replace verbose null checks with `?.`: `user?.profile?.address?.zipCode`.
  - Use `??` (nullish coalescing) when falling back from `null` or `undefined`, preserving falsy valid values like `0` or `false` (avoid `||` when `0` is a valid number):
    ```javascript
    const count = rawCount ?? 0;
    const isEnabled = config.enabled ?? true;
    ```
- **Template Literals**: Avoid string concatenation (`+`) for dynamic strings. Use clean template literals with `${expression}`.
- **Data Structures**: Leverage `Set` for deduplication and $O(1)$ lookups, and `Map` or plain keyed objects for dictionary associations instead of parallel arrays.

---

## 2. Declarative Data Processing

Avoid imperative `for (let i = 0; i < len; i++)` loops for array transformations unless index mutation is required. Use functional array methods:
- **`map`**: Transform every element into a new structure.
- **`filter`**: Select elements meeting a predicate.
- **`reduce`**: Aggregate arrays into a single value, dictionary, or grouped map.
- **`some` / `every`**: Test conditions across elements with short-circuiting.
- **`find` / `findIndex`**: Search elements without iterating entire collections.

```javascript
// Clean declarative grouping:
const debtsByPerson = debtItems.reduce((acc, item) => {
  const person = item.person;
  acc[person] = (acc[person] || 0) + item.amount;
  return acc;
}, {});
```

---

## 3. Asynchronous & Promise Handling

- **`async` / `await` Syntax**: Always prefer `async`/`await` over chained `.then().catch()` callbacks.
- **Error Boundaries**: Every `await` operation that can fail (network requests, filesystem, external APIs) must be enclosed within a meaningful `try/catch` block or return a structured result tuple.
- **Parallel Execution**: Execute independent asynchronous operations concurrently with `Promise.all()` or `Promise.allSettled()`:
  ```javascript
  const [userData, ordersData] = await Promise.all([
    fetchUserProfile(userId),
    fetchUserOrders(userId)
  ]);
  ```
- **No Unhandled Rejections**: Ensure all asynchronous promises have an explicit rejection handler.

---

## 4. Defensive Programming & Safety

- **Date Parsing**: Never rely on raw `new Date("some-string")` across different locales. Always validate timestamp integrity with `!isNaN(date.getTime())` and parse structured date components (e.g. ISO `YYYY-MM-DD` or `DD/MM/YYYY`) explicitly.
- **Number Parsing**:
  - Clean strings of currency symbols and commas before parsing: `parseFloat(val.replace(/[^\d.-]/g, ""))`.
  - Always guard with `!isNaN(amount)` and handle zero/negative limits explicitly.
- **Fail-Safe Property Access**:
  - Guard JSON serialization and deserialization with `try/catch` blocks.
  - Provide fallback defaults when reading configuration or environment properties.

---

## 5. Type Safety & JSDoc Annotations

In JavaScript projects without full TypeScript compilation (such as direct Apps Script projects), enforce type clarity and IDE autocomplete using comprehensive JSDoc comments:

```javascript
/**
 * Represents an itemized debt record for a family member.
 * @typedef {Object} DebtRecord
 * @property {string} person - The normalized name of the debtor.
 * @property {number} amount - The owed amount in EGP.
 * @property {Date} dueDate - The scheduled due date of the billing cycle.
 * @property {boolean} isSettled - Whether the debt has been marked paid.
 */

/**
 * Calculates the total outstanding balance across all debtors.
 * @param {DebtRecord[]} records - Array of individual debt entries.
 * @returns {number} The aggregate sum of outstanding debts.
 */
function calculateOutstandingTotal(records) {
  if (!Array.isArray(records)) return 0;
  return records
    .filter(r => !r.isSettled)
    .reduce((sum, r) => sum + r.amount, 0);
}
```

---

## 6. Code Structure & Anti-Patterns to Avoid

| Anti-Pattern | Why It Fails | Modern Replacement |
| :--- | :--- | :--- |
| `var` declarations | Leaks function scope, hoists unintentionally | `const` (immutable ref) and `let` (reassignable) |
| Global mutations | Unpredictable side effects, difficult testing | Pure functions with explicit parameters and return values |
| Deep callback nesting | "Pyramid of Doom", unreadable error flow | `async` / `await` with early returns |
| Loose equality (`==`) | Unintended type coercion (`0 == false`, `"" == 0`) | Strict equality (`===` and `!==`) |
| Silent error swallowing | `catch (e) {}` conceals breaking bugs | Structured logging (`Logger.log`, `console.error`) with user alerts |
