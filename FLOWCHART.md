# 🏗️ TESDA ID Generator - System Architecture & Logic

This document provides a comprehensive overview of how the ID Generator operates, from user interaction to the internal document transformation logic.

---

## 🗺️ High-Level Application Flow

The following diagram breaks down the lifecycle of an ID generation session into four distinct phases: **Setup**, **Ingestion**, **Transformation**, and **Output**.

```mermaid
graph TD
    %% Phase 1: Setup
    subgraph Phase_1 [1. Setup & Detection]
        Start([Start App]) --> LoadDoc[/Upload .docx Template/]
        LoadDoc --> Validate{Valid Template?}
        Validate -- No --> LoadDoc
        Validate -- Yes --> Detect[[Auto-detect Placeholders]]
    end

    %% Phase 2: Ingestion
    subgraph Phase_2 [2. Data Ingestion]
        Detect --> InputType{Select Input Method}
        InputType -- "Manual Entry" --> TextArea[Type names/data in UI]
        InputType -- "Bulk Upload" --> CSV[(Read CSV File)]
        TextArea --> Parse[Parse Lines & Map to Fields]
        CSV --> Parse
    end

    %% Phase 3: Transformation
    subgraph Phase_3 [3. Document Transformation]
        Parse --> Loop[Iterate through Placeholders]
        Loop --> Match{Match Found?}
        Match -- "Name Field" --> Upper[<b>Convert to ALL CAPS</b>]
        Match -- "Other Fields" --> Raw[Keep Original Case]
        Upper --> Replace[Update XML Text Node]
        Raw --> Replace
        Replace --> Next{More Nodes?}
        Next -- Yes --> Loop
    end

    %% Phase 4: Output
    subgraph Phase_4 [4. Finalization]
        Next -- No --> Save[(Save Timestamped .docx)]
        Save --> OpenDir[Open Output Folder]
        OpenDir --> End([End Session])
    end

    %% Styling
    style Upper fill:#ff9f43,stroke:#333,stroke-width:3px,color:#fff
    style Phase_1 fill:#f0f7ff,stroke:#0056b3
    style Phase_2 fill:#fff9eb,stroke:#d4a017
    style Phase_3 fill:#f0fff4,stroke:#28a745
    style Phase_4 fill:#fff5f5,stroke:#c53030
```

---

## 📖 Detailed Logic Explanation

### 1. Setup & Detection
*   **Template Analysis**: The app doesn't just "write" over a file; it parses the internal XML structure of the `.docx`.
*   **Auto-Detection**: It specifically looks for defined keys like `NAME HERE` or `CODE HERE`. If your template has two IDs per page, it intelligently identifies duplicate placeholders to ensure both are updated for the same person.

### 2. Data Ingestion
*   **Flexible Mapping**: The system supports multiple CSV formats (1, 4, 8, or 9 columns). 
*   **Smart Parsing**: Whether you paste data with tabs, commas, or multiple spaces, the backend normalizes the input to prevent "empty" data from being generated.

### 3. Transformation (The "Brain")
*   **Case Normalization**: To ensure professional standards, the **Name** field is programmatically intercepted. Even if a user types `john doe`, the system applies `.upper()` during the XML injection phase to produce `JOHN DOE`.
*   **XML Injection**: The app reaches into the document's `word/document.xml` and replaces the text within `<w:t>` tags without breaking the surrounding formatting (font, size, color).

### 4. Finalization
*   **Non-Destructive Saving**: To protect your original template, the system generates a unique filename using the pattern: `UPDATED_IDS_YYYYMMDD_HHMMSS.docx`.

---

## 🎨 Team Design Standards (Legend)

Use these shapes and colors when expanding this documentation:

| Shape | Role | Mermaid Code |
| :--- | :--- | :--- |
| **Rounded Oval** | Start/End Points | `([Text])` |
| **Rectangle** | Standard Action | `[Text]` |
| **Rhombus/Diamond** | Logic Decision | `{Text}` |
| **Parallelogram** | User Input/File Upload | `[/Text/]` |
| **Cylinder** | Database or Saved File | `[(Text)]` |
| **Double-Border** | Complex Sub-Process | `[[Text]]` |

### **Color Coding Scheme**
- 🔵 **Blue**: Setup & UI Initialization
- 🟡 **Yellow**: Data Handling & Parsing
- 🟢 **Green**: Internal Logic & Transformation
- 🔴 **Red**: File Output & System Completion
