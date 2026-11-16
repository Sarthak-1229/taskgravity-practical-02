# Practical 02: Interactive Web Application using JavaScript

## TaskGravity — JSON & Array Data Model, Client-Side Validation, and Dynamic DOM Manipulation

### 1. Overview
Practical 02 transitions the TaskGravity Task Management System from static presentation into an interactive client-side application using standard HTML5, CSS3, and modern Vanilla JavaScript (ES6+).

### 2. Key Features
- **JSON & Array Data Persistence**: Tasks are modeled as JSON objects stored within an in-memory array and synchronized with browser `localStorage`.
- **Comprehensive Input Validation**:
  - `title`: Required, length 3-80 chars.
  - `category`: Must match valid architectural categories (`Frontend`, `Backend`, `Database`, `DevOps`, `Testing`, `UI/UX`).
  - `priority`: Must match valid priority levels (`Low`, `Medium`, `High`, `Urgent`).
  - `dueDate`: Required, must be a valid future or present date.
  - `estimatedHours`: Numeric, strictly between 1 and 100 hours.
  - `description`: Optional, capped at 300 characters.
- **Dynamic Event-Driven UI**:
  - Real-time input validation on keystroke (`input`) and blur (`change`).
  - Add task with field-specific error feedback badges.
  - Cycle task status (`Pending` -> `In Progress` -> `Completed`) via checkbox or action buttons.
  - Delete task with user confirmation.
  - Live full-text search across task titles and descriptions.
  - Multi-attribute filtering (by status, category) and sorting (by due date, priority, effort, title).
  - Dynamic KPI / Metric calculation (Total, In Progress, Pending Review, Velocity Percentage).
  - Floating toast notifications.

### 3. Execution
Simply open `index.html` in any modern web browser or serve locally:
```bash
npx serve .
```
