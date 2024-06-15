# Ansible Collection: yumizu11.excel

The ```yumizu11.excel``` Ansible Collection includes the module to read Excel cells and also the module to write list of data to Excel file.

## Ansible version compatibility

This collection has been tested against the following Ansible version: >= 2.9.10

## Installation and Usage

### Installing the Collection from Ansible Galaxy

```
ansible-galaxy collection install yumizu11.excel
```

You can also include it in a ```requirements.yml``` file and install it via ```ansible-galaxy collection install -r requirements.yml``` using the format:

```yaml
collections:
  - name: yumizu11.excel
```

## Modules

This collection provides the following modules you can use in your own roles and playbooks:

|Name|Description|
|---|---|
|read_sheet|Read Excel file and return specified range of cell values|
|write_sheet|Write list of values to the specified sheet in the specified Excel file|

### Examples

#### read_sheet module:

```
# Read a whole rows of the 1st sheet
- name: Read whole rows
  yumizu11.excel.read_sheet:
    path: mysheet.xlsx
  register: sheet_rows

# Read the specified sheet
- name: Read whole rows of mysheet
  yumizu11.excel.read_sheet:
    path: mysheet.xlsx
    sheet: mysheet
  register: mysheet_rows

# Read specified range of cells
- name: Read B2 to D5
  yumizu11.excel.read_sheet:
    path: mysheet.xlsx
    startrow: 2
    endrow: 5
    startcolumn: 2
    endcolumn: 5
  register: sheet_b2_d5

# Read specified range of cells
- name: Read sheet with header in the 2nd row
  yumizu11.excel.read_sheet:
    path: mysheet.xlsx
    startrow: 2
    header: true
  register: sheet_rows
```

#### write_sheet module:

```
# Write 2D array to the first sheet of mysheet.xlsx
- name: Write whole data
  vars:
    mydata:
      - [1,2,3]
      - [4,5,6]
  yumizu11.excel.write_sheet:
    path: mysheet.xlsx
    data: "{{ mydata }}"

# Write a list of map to sheet 'sheetA'
- name: Write whole data
  vars:
    mydata:
      - firstname: Taro
        lastname: Yamada
      - firstname: Hanako
        lastname: Sato
  yumizu11.excel.write_sheet:
    path: mysheet.xlsx
    sheet: sheetA
    data: "{{ mydata }}"

# Write 2D array to cell C3~ on the first sheet
- name: Write whole data
  vars:
    mydata:
      - [1,2,3]
      - [4,5,6]
  yumizu11.excel.write_sheet:
    path: mysheet.xlsx
    data: "{{ mydata }}"
    startrow: 3
    startcolumn: 3

# Insert a list of map on the cell C3
- name: Write whole data
  vars:
    mydata:
      - firstname: Taro
        lastname: Yamada
      - firstname: Hanako
        lastname: Sato
  yumizu11.excel.write_sheet:
    path: mysheet.xlsx
    sheet: sheetA
    data: "{{ mydata }}"
    startrow: 3
    startcolumn: 3
    insert: True
```

## Plugins

This collection provides the following filters you can use in your own roles and playbooks:

|Name|Description|
|---|---|
|read_excel|Read whole sheet and return 2 dimension list of cell values|
|read_excel_range|Read specified range of sheet and return 2 dimension list of cell values|
|read_excel_header|Read whole sheet which has header row on the top row and return a list of map|
|read_excel_header_rage|Read specified range of sheet which has header row on the top row and return a list of map|

### Examples

#### read_excel:

```
- name: Show whole cell values of the 1st sheet of sheet1.xlsx
  vars:
    excel_path: /tmp/sheet1.xlsx
  ansible.builtin.debug:
    msg: "{{ excel_path | read_excel }}"

- name: Show whole cell values of sheet 'sheet1' of sheet1.xlsx
  vars:
    excel_path: /tmp/sheet1.xlsx
  ansible.builtin.debug:
    msg: "{{ excel_path | read_excel('sheet1') }}"
```

#### read_excel_range:

```
- name: Show values of cell B2 to E10 of the 1st sheet of sheet1.xlsx
  vars:
    excel_path: /tmp/sheet1.xlsx
    startrow: 2
    startcolumn: 2
    endrow: 10
    endcolumn: 5
  ansible.builtin.debug:
    msg: "{{ excel_path_range | read_excel(startrow, startcolumn, endrow, endcolumn) }}"

- name: Show values of cell B2 to E10 of the sheet 'sheet1' of sheet1.xlsx
  vars:
    excel_path: /tmp/sheet1.xlsx
    startrow: 2
    startcolumn: 2
    endrow: 10
    endcolumn: 5
  ansible.builtin.debug:
    msg: "{{ excel_path_range | read_excel(startrow, startcolumn, endrow, endcolumn, 'sheet1') }}"
```

#### read_excel_header:

```
- name: Show list of map (cell item to value) of the 1st sheet of sheet1.xlsx
  vars:
    excel_path: /tmp/sheet1.xlsx
  ansible.builtin.debug:
    msg: "{{ excel_path | read_excel_hedarer }}"

- name: Show list of map (cell item to value)  of sheet 'sheet1' of sheet1.xlsx
  vars:
    excel_path: /tmp/sheet1.xlsx
  ansible.builtin.debug:
    msg: "{{ excel_path | read_excel('sheet1') }}"
```

#### read_excel_range:

```
- name: Show list of map (cell item to value) of cell B2 to E10 of the 1st sheet of sheet1.xlsx
  vars:
    excel_path: /tmp/sheet1.xlsx
    startrow: 2
    startcolumn: 2
    endrow: 10
    endcolumn: 5
  ansible.builtin.debug:
    msg: "{{ excel_path_range | read_excel(startrow, startcolumn, endrow, endcolumn) }}"

- name: Show list of map (cell item to value) of cell B2 to E10 of sheet 'sheet1' of sheet1.xlsx
  vars:
    excel_path: /tmp/sheet1.xlsx
    startrow: 2
    startcolumn: 2
    endrow: 10
    endcolumn: 5
  ansible.builtin.debug:
    msg: "{{ excel_path_range | read_excel(startrow, startcolumn, endrow, endcolumn, 'sheet1') }}"
```

## License

MIT License

See LICENSE.txt to see full text.
