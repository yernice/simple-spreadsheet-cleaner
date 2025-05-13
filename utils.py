# Decides if we should skip the current row
def skip(sheet, row):
    for col in range(1, sheet.max_column + 1):
        if sheet.cell(row=row, column=col).value is None:
            return True

    return False


# Cleans up the value
def clean_value(value):
    if isinstance(value, str):
        value = value.strip()
        if '@' in value:
            value = value.lower()
        else:
            value = value.title()

    return value


# Copies the header row
def copy_first_row(input, output):
    for col in range(1, input.max_column + 1):
        output.cell(row=1, column=col).value = input.cell(
            row=1, column=col).value


# Checks if the current row is duplicate
def duplicate(input, output, in_row, out_row):

    in_name = input.cell(row=in_row, column=1).value
    in_age = input.cell(row=in_row, column=3).value
    in_name = clean_value(in_name)

    for row in range(2, out_row):
        out_name = output.cell(row=row, column=1).value
        out_age = output.cell(row=row, column=3).value

        if in_name == out_name and in_age == in_age:
            return True

    return False
