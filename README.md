# ConsolidatedShortageTable-GoogleAppsScript
## Aim:
- Fetch Batch Q:AP
- Single Active Cell Logging
- Column F Parsing 
- Daily / Weekly / Overhead Matching 
- Adjusted-Current Calculation 

### Overall Flow
```
Read:
    Current Inventory Value / Balance for each date range
    +
    Previous Inventory Values
    +
    Initial Stock ( L + M+ N)
    +
    Column F planned quantities
    
↓
Calculate Movement (xValue)
↓
Write one record into Movements:
(Date, Item No, Movement)

Current = Current + Column F Qty
Movement = Current - Previous
```

## VBA Usage Setup Instrutions



### Windows (Excel Desktop)

1. Open the attached `.xlsm` file.
2. If a security warning appears at the top of the screen, click **Enable Editing**, and then click **Enable Content**.
3. Press **Alt + F8** to open the Macro list.
4. Run either of the following macros:
   * **LogMovement**: Calculates the currently selected single cell.
   * **LogAllMovements**: Batch calculates all rows (starting from Row 5) for columns Q to AP, and writes the results into the `Movements` worksheet.

---

### Mac (Excel Desktop)

1. Open the `.xlsm` file.
2. If the system prompts you to enable macros, select **Enable Macros**.
3. Click **Tools** → **Macro** → **Macros** (or **Developer** → **Macros**, depending on your Excel version).
4. Run **LogMovement** or **LogAllMovements**.

---

> [!IMPORTANT]
> **Reminders:**
> * Please use Microsoft Excel Desktop to open the file; the web version of Excel cannot run VBA macros.
> * Please keep the file in `.xlsm` format. If you save it as `.xlsx`, the macro code will be removed.
