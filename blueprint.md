
# Spreadsheet Application Blueprint

## Overview

This document outlines the architecture and implementation of a spreadsheet application built with React. The application provides a dynamic grid interface for data manipulation, including adding/deleting rows and columns, and saving/loading data to and from local files.

## Project Structure

The project is initialized as a standard Vite + React application with the following key files:

- `index.html`: The main HTML entry point.
- `src/main.jsx`: The main React rendering entry point.
- `src/App.jsx`: The root component of the application.
- `src/App.css`: Styles for the `App` component.
- `src/index.css`: Global styles.
- `src/ErrorBoundary.jsx`: A component to catch and display runtime errors.

## Core Features & Design

- **Dynamic Grid:** The core of the application is a data grid built using `@mui/x-data-grid`, allowing for editable cells and row selection.
- **Row & Column Manipulation:** Users can add new rows and columns, as well as delete selected rows and columns.
- **Modern Styling:** The application uses `@mui/material` for UI components and a clean, modern aesthetic.
- **Error Handling:** A custom `ErrorBoundary` component is implemented to prevent the entire application from crashing due to runtime errors.
- **File-based Save/Load:** Users can save the state of their spreadsheet to a local JSON file and load it back into the application.

## Recent Changes

### Consolidated Actions Menu

- **Requirement:** Clean up the cluttered header by consolidating the various action buttons into a single menu.
- **Solution:** Replaced the individual buttons for "Save," "Load," "Add Row/Column," "Insert Row/Column," and "Delete Row/Column" with a single "Actions" button. This button opens a dropdown menu containing all the previous actions, which are logically grouped using dividers for improved usability.

### Column and Row Resizing

- **Requirement:** Allow users to resize the width of columns and the height of rows.
- **Solution:**
    - **Column Resizing:** This is a built-in feature of the underlying `@mui/x-data-grid` component and is enabled by default. Users can resize columns by dragging the border between column headers.
    - **Automatic Row Height:** Implemented a feature where rows automatically adjust their height to fit their content. This provides similar benefits to manual resizing and ensures that all data within a cell is visible.

### Robust Formula Engine & Data Integrity

- **Requirement:** Fix critical bugs that caused formulas to break (`#REF!`) when rows or columns were inserted or deleted.
- **Solution:** Re-architected the core data and formula management system.
    - **Decoupled Logic:** The formula engine is no longer dependent on the visual position (array index) of a row. It now uses a stable, unique ID for each row, combined with a positional map for lookups.
    - **Context-Aware Formula Rewriting:** When a row or column is inserted or deleted, the application now parses and rebuilds all affected formulas with updated, correct cell references. This is a robust, context-aware process that prevents reference errors.
    - **Reliable Deletion:** Deleting rows or columns that are referenced by a formula will now correctly result in a `#REF!` error in the dependent cell, as expected.

### Relative Row and Column Insertion

- **Requirement:** Allow users to insert new rows and columns at specific positions, not just at the end of the grid.
- **Solution:**
    - Added four new buttons: "Insert Row Above," "Insert Row Below," "Insert Column Left," and "Insert Column Right."
    - These buttons are enabled only when a cell is selected, using the active cell as the reference point for the insertion.
    - Implemented logic to shift existing rows/columns and automatically adjust all formula references to maintain data integrity.

### Interactive Formula Bar

- **Requirement:** Allow users to view and edit the formula of a selected cell in a dedicated input field.
- **Solution:** 
    - Implemented a `TextField` component above the grid to serve as a formula bar.
    - The bar is synchronized with the currently selected cell (`cellSelectionModel`).
    - When a cell is selected, its raw content is displayed in the formula bar.
    - Users can now edit the content in the formula bar and press **Enter** to commit the change to the selected cell, which triggers a full recalculation of the spreadsheet.

### Expanded Formula Engine

- **Requirement:** Add support for more advanced spreadsheet functions.
- **Solution:** 
    - Refactored the formula engine to be more extensible.
    - Implemented `SUM`, `AVERAGE` (aliased as `MEAN`), and `MEDIAN` functions.
    - These functions operate on cell ranges (e.g., `=SUM(A1:B10)`).
    - Added proper error handling for cases like division by zero (`#DIV/0!`) or invalid input (`#NUM!`).
