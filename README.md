# excelFuzzyMatcherLearnable
 ---
VBA/Power Query-based solution for automated fuzzy text matching:
- **VBA macros** (event processing, user interface)
- **Power Query** (ETL processes, fuzzy-merge)
- **User-defined functions** (Levenshtein distance calculation)
---
## Prerequisites
- **Excel 2019+**
- **Macros enabled**: `File → Options → Trust Center → Enable all macros`
- **Power Query enabled**: `File → Options → Add-ins → Manage COM Add-ins → Power Query`

##  Setup & Workflow

### 1. Data Preparation
- **Database** (`db/` folder):
  - Add CSV files with columns: `code`, `name`, `price`, `comment` other columns is up to user
    ```csv
    code,name,price,comment
    103-024,Respiratory specimens,160,Sputum
    ```
- **Search queries** (`search/` folder):
  - Add XLSX files with search terms in **Column A** (Sheet1):
    ```
    Respiratory screen
    ```
### 2. Initialization
1. Open `searching.xlsm`
2. **Macro auto-actions**:
   - Sets `settingsMain[value] = path` (current file path)

### 3. Data Processing Flow & Manual Intervention:
   - Refresh workflow (`Ctrl+Alt+F5` or `Data` → `Refresh All`)
   - Power Query executes in this order:
      - Loads latest files from:
        - search/ (XLSX) → search table
        - db/ (CSV/XLSX) → db table
   - Merge search data with your db datas with `settingsMain[value] = fuzzyThreshold` fuzzy treshold atribute setted(0,7 by default), and `fromTo`  transformation table 
   - Output format:
     
    | search                        | code    | name                    | price | comment       | levenstein |
    |-------------------------------|---------|-------------------------|-------|---------------|------------|
    | Respiratory screen            |103-024  |Respiratory specimens    |160    |Sputum         | 18         |

  **Manual Matching via VBA Interface**:
  - **Double-click** any empty cell in the `code` column → opens search window:
    ![image](https://github.com/user-attachments/assets/40db2bd0-3655-42ec-b50f-b99512c329b8)
## Search Workflow
| Step | Action | Result |
|------|--------|--------|
| 1 | Type in search box | Real-time filtering of `db` records |
| 2 | Select match (Double-click/Enter) | `code` inserted, other fields auto-populated |

### 4. Update FromTo transformation table
 - After manual corrections**, click the `UpdateFromTo` button on the "searching" sheet, trigger `updateFromTo` VBA script:
 	 1. Scan all rows in the `searching` table
 	 2. Identify entries where `levenstein <5`
 	 3. Add `search - name` pairs to `fromTo` tabble 
