import operator as op
import ast

# Custom Errors
class CustomSheetIndexError(IndexError):
	def __init__(self, index):
		"""Raised for an invalid cell index."""
		super().__init__(f"The provided cell index '{index}' is not valid.")

class CustomEmptyCellError(Exception):
	def __init__(self, index):
		"""Raised when attempting to edit or delete a value in a cell that has not been initialized."""
		super().__init__(f"The cell at '{index}' has not been initialized yet. Use set_value() to initialize it.")

class CustomValueAlreadySetError(Exception):
	def __init__(self, index, value):
		"""Raised when attempting to set a value in a cell that already has a value."""
		super().__init__(f"The cell at '{index}' already has a value ({value}). Use edit_value() to change existing values.")


class Sheet():
	_data_array = []  # 2D array to store cell values
	_alphabet_mapping = {}  # Mapping of column letters to indices (A -> 0, B -> 1, ..., Z -> 25)

	_operators = {
		ast.Add: op.add,       # Addition operator
		ast.Sub: op.sub,       # Subtraction operator
		ast.Mult: op.mul,      # Multiplication operator
		ast.Div: op.truediv,   # Division operator
		ast.USub: op.neg,      # Unary negation operator
		ast.Mod: op.mod,       # Modulo operator
		ast.Pow: op.pow        # Exponentiation operator
	}

	def __init__(self, default_value=None):
		"""
		Initialize the sheet, creating an alphabet mapping for column indices.
		"""
		self.default_value = default_value
		# Create a dictionary
		self._alphabet_mapping = {chr(char_code): i for i, char_code in enumerate(range(ord('A'), ord('Z') + 1))}

	def __resolve_index(self, str_index: str) -> (int, int):
		"""
		Convert the cell index from a string format (e.g., "A1") to row and column indices.
		"""
		str_index = str_index.upper()
		head = str_index.rstrip('0123456789')
		tail = str_index[len(head):]

		column_index = self._alphabet_mapping.get(head)
		if column_index is None:
			raise CustomSheetIndexError(str_index)

		row_index = int(tail) - 1  # Numeration of cells begins from 1
		return row_index, column_index

	def __evaluate_formula(self, formula: str):
		"""
		Evaluate a formula by parsing it into an abstract syntax tree and then using the __eval method.
		"""
		formula = formula[1:]  # Remove the '=' prefix

		try:
			ast_tree = ast.parse(formula, mode='eval').body
			result = self.__eval(ast_tree)
		except (RecursionError, TypeError):
			result = "ERROR"

		return result

	def __eval(self, node):
		"""
		Recursively evaluate mathematical expressions represented as abstract syntax trees.
		"""
		match node:
			case ast.Name(node):
				# Cell ID
				return self.get_value(node, raw=True)
			case ast.Constant(value) if isinstance(value, int) or isinstance(value, float):
				# Number
				return value
			case ast.BinOp(left, op, right):
				# Binary operation
				left = self.__eval(left)
				right = self.__eval(right)
				op_result = self._operators[type(op)](left, right)
				return op_result
			case ast.UnaryOp(op, operand):
				# Unary operation
				return self._operators[type(op)](self.__eval(operand))
			case _:
				raise TypeError(node)

	def resize(self, rows, cols):
		"""
		Resize the sheet to the specified number of rows and columns.
		"""
		current_height = len(self._data_array)
		current_width = len(self._data_array[0]) if self._data_array else 0

		# Expand or shrink height
		self._data_array = self._data_array[:rows] + [[self.default_value] * cols for _ in range(rows - current_height)]

		# Expand or shrink width
		for sub_arr in self._data_array:
			sub_arr.extend([self.default_value] * (cols - current_width) if cols > current_width else [])

	# Value Manipulation Methods (set, get, edit, delete)
	def set_value(self, str_index, value):
		"""
		Set a value in the specified cell, handling index errors and checking if the cell already has a value.
		"""
		row, col = self.__resolve_index(str_index) # Convert the cell index

		try:
			# Attempt to retrieve the current value in the cell
			current_value = self._data_array[row][col]
		except IndexError:
			# If the cell doesn't exist, resize the sheet and set the value
			self.resize(row + 1, col + 1)
			self._data_array[row][col] = value
		else:
			# If the cell already has a value, check and handle accordingly
			if current_value != self.default_value:
				raise CustomValueAlreadySetError(str_index, current_value)
			self._data_array[row][col] = value

	def get_value(self, str_index, raw=False):
		"""
		Get the value from the specified cell, evaluating formulas if present and returning both the raw value and the result.
		"""
		row, col = self.__resolve_index(str_index) # Convert the cell index
		value = self._data_array[row][col] # Retrieve the raw value from the specified cell
		result = value # Initialize result with the raw value
		
		# Check if the value is a formula
		if isinstance(value, str) and value.startswith('='):
			result = self.__evaluate_formula(value) # Evaluate the formula and update the result

		# Return the raw value and result based on the 'raw' parameter
		if not raw:
			return {"value": value, "result": result}
		elif raw:
			return result

	def edit_value(self, str_index, value):
		"""
		Edit the value in the specified cell, handling empty cell errors.
		"""
		row, col = self.__resolve_index(str_index) # Convert the cell index
  
		try:
			# Attempt to edit the value in the cell
			self._data_array[row][col] = value
		except IndexError:
			# If the cell doesn't exist, raise a CustomEmptyCellError
			raise CustomEmptyCellError(str_index)

	def delete_value(self, str_index):
		"""
		Delete the value in the specified cell, handling empty cell errors.
		"""
		row, col = self.__resolve_index(str_index) # Convert the cell index
  
		try:
			# Attempt to delete the value in the cell
			self._data_array[row][col] = self.default_value
		except IndexError:
			# If the cell doesn't exist, raise a CustomEmptyCellError
			raise CustomEmptyCellError(str_index)

	# Magic Methods (__str__, __setitem__, __getitem__)
	def __str__(self):
		"""
		Convert the sheet to a string representation, mainly for debugging purposes.
		"""
		return '\n'.join(map(str, self._data_array))

	def __setitem__(self, key, value):
		"""
		Allows setting values using the indexing syntax (sheet[key] = value).
		"""
		self.set_value(self, key, value)

	def __getitem__(self, key):
		"""
		Allows getting values using the indexing syntax (sheet[key]).
		"""
		return self.get_value(key)
