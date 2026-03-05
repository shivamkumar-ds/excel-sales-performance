## 1. Data Sorting
Sorting allows you to reorder data alphabetically, numerically, or by date to identify trends and outliers.

### Quick Sort
* **Location:** Home Tab > Editing > Sort & Filter.
* **Ascending (A-Z / Smallest to Largest):** Groups data from lowest to highest.
* **Descending (Z-A / Largest to Smallest):** Groups data from highest to lowest.
* **Warning:** Ensure your cursor is inside the data set before sorting. If you select only one column and sort, Excel will prompt a "Sort Warning." Always choose **"Expand the selection"** to keep row data together.

### Multi-Level Sort
Use this when you need to sort by more than one criteria (e.g., Sort by "Department" first, then by "Last Name" within that department).
1. Go to **Data Tab > Sort**.
2. Click **Add Level** to define the secondary sorting rule.

---

## 2. Data Filtering
Filtering hides data that doesn't meet specific criteria, allowing you to focus on a subset of information without deleting anything.

### Applying Filters
* **Shortcut:** Select any cell in your header row and press `Ctrl + Shift + L`.
* **The Filter Drop-down:** Click the arrow in the header to:
    * **Text Filters:** Contains, Begins with, or specific names.
    * **Number Filters:** Greater than, Less than, or Top 10.
    * **Color Filters:** Filter by cell background or font color.



---

## 3. Conditional Formatting
This feature automatically applies formatting (colors, icons, etc.) to cells only when they meet a specific condition. It makes data "visual."

### Highlight Cells Rules
* **Greater Than / Less Than:** Highlights values above or below a threshold (e.g., highlighting all sales > $5,000 in green).
* **Text that Contains:** Highlights specific words (e.g., highlighting "Overdue" in red).
* **Duplicate Values:** Quickly finds and highlights repeated entries in a list.

### Visual Data Tools
* **Data Bars:** Adds a horizontal bar inside the cell. Longer bars represent higher values.
* **Color Scales:** Applies a gradient (e.g., Red-Yellow-Green) where colors represent the value's position in the range.
* **Icon Sets:** Adds symbols like green checkmarks, yellow exclamation points, or red "X"s based on the cell's value.



### Managing Rules
If you need to change a color or a threshold, go to **Conditional Formatting > Manage Rules**. You can edit or delete existing logic from here.

---

## 4. Professional Best Practices
* **Clear Filters First:** Before performing a new analysis, ensure no hidden filters are active. Check if the row numbers are blue (indicating a filter is on).
* **Use Clear Colors:** When using Color Scales, ensure the contrast is high enough to read the numbers behind the color.
* **Filter by Search:** Use the search box inside the filter drop-down for large datasets to find a specific record instantly.

---

## Key Takeaways
* **Sorting** changes the order; **Filtering** changes the visibility.
* **Conditional Formatting** is dynamic—if the data changes, the formatting updates automatically.
* **Expand Selection:** Always ensure the whole table is included when sorting to avoid "scrambling" your data.

## Common Errors
* **Partial Sorting:** Sorting one column without expanding the selection, which breaks the relationship between rows (e.g., a name gets paired with the wrong phone number).
* **Hidden Data:** Forgetting a filter is active and thinking data has been deleted.
* **Rule Overlap:** Applying multiple Conditional Formatting rules to the same range can cause visual clutter. Use "Manage Rules" to prioritize them.