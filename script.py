from sheet import Sheet

sh = Sheet()

# Set values to B1 and B2
sh.set_value("B1", 1)
sh.set_value("B2", 2)

print(sh.get_value('B1')) # => {'value': 1, 'result': 1}
print(sh.get_value('B2')) # => {'value': 2, 'result': 2}

# Calculate B3 based on B1 and B2
sh.set_value("B3", '=B1+B2') 
print(sh.get_value('B3')) # => {'value': '=B1+B2', 'result': 3}

# Calculate cell based on calculated cell
sh.set_value("B4", '=B3-10')
print(sh.get_value('B4')) # => {'value': '=B3-10', 'result': -7}

# Set value to different column
sh.set_value("A1", 10)
print(sh.get_value('A1')) # => {'value': 10, 'result': 10}

# Refer to another column
sh.set_value("A2", '=B1*2')
print(sh.get_value('A2')) #=> {'value': '=B1*2', 'result': 2}
sh.set_value("A3", '=A2/2')
print(sh.get_value('A3')) # => {'value': '=A2/2', 'result': 1.0}

# Edit refered value
sh.edit_value('B2', 20)
print(sh.get_value('B3')) # => {'value': '=B1+B2', 'result': 21}

# Delete reference
sh.delete_value('B1')
print(sh.get_value('B3')) # => {'value': '=B1+B2', 'result': 'ERROR'}
