# Spreadsheet Python
https://github.com/povstenko/sheet-python
## Description
This project introduces a `Sheet` class, a simple spreadsheet implementation in Python. The `Sheet` class provides functionality to manipulate data in a tabular format. Users can interact with the sheet by setting, retrieving, editing, and deleting values in specific cells. Additionally, the class supports basic formula evaluation for dynamic cell calculations.

## Files
- `sheet.py`: Contains the implementation of the `Sheet` class.
- `script.py`: A script demonstrating the usage of the `Sheet` class.

## How to Run
Execute the `script.py` file using the following command:

```bash
python script.py
```

## `Sheet` Class
The `Sheet` class is the core component of this project, offering a basic spreadsheet functionality. Below are key features and methods of the `Sheet` class:

### Initialization
```python
from sheet import Sheet

# Create a new Sheet instance
sh = Sheet()
```
### Cell Indexing
Cells are identified using a combination of column letters and row numbers (e.g., "A1", "B2"). The class provides methods to convert these string indices into row and column indices.

```python
# Convert cell index from string to row and column indices
row, col = sh.__resolve_index("B3")
```
### Value Manipulation
The `Sheet` class provides methods to set, get, edit, and delete values in cells.

```python
# Set a value in a specific cell
sh.set_value("A1", 42)

# Retrieve the value from a cell
value = sh.get_value("A1")

# Edit the value in a cell
sh.edit_value("A1", 100)

# Delete the value in a cell
sh.delete_value("A1")
```
### Formula Evaluation
Formulas can be assigned to cells, allowing dynamic calculations based on other cell values.

```python
# Set a formula in a cell
sh.set_value("B3", '=A1 + A2')

# Retrieve the result of a formula
result = sh.get_value("B3")
```
### Resizing
The `resize` method allows resizing the sheet to a specified number of rows and columns.

```python
# Resize the sheet to 5 rows and 5 columns
sh.resize(5, 5)
```
## Examples
The `script.py` file provides examples demonstrating the usage of the `Sheet` class, including setting values, performing calculations with formulas, and more.

```python
# Example usage in script.py
from sheet import Sheet

sh = Sheet()

sh.set_value("A1", 10)
sh.set_value("B1", 20)

result = sh.get_value("A1")  # {'value': 10, 'result': 10}

sh.set_value("C1", '=A1 + B1')

result = sh.get_value("C1")  # {'value': '=A1 + B1', 'result': 30}
```
## License

## Contributing

## Contact Information
