Project Name
============

This is a Python project for analyzing CSV files that contains HVAC project. The project reads a CSV file, performs analysis, and saves the results in a new CSV(excel) file.

Folder Structure
----------------

```lua
project_name/
├── data/
│   └── input/
│       └── input.csv
│   └── output/
│       └── output.xlsx
├── scripts/
│   └── analyze_csv.py
├── requirements.txt
└── README.md
```

 Dependencies
 ------------

 * pandas
 * numpy
 * openpyxl

 Input
 -----

 The script reads a CSV file containing procurement data for a specific item. The path to the file is specified in `input_file_path`. The expected format of the CSV file is:
 
``` 
  
| 계약(납품요구)일자 | 계약(납품요구)번호 | 수요기관구분 | 수요기관지역명 | 물품분류번호 | 품명 | 세부물품분류번호 | 물품식별번호 | 세부품명 | 품목 | 단가 | 수량 | 단위 | 금액 | |------------------|------------------|--------------|----------------|--------------|------|----------------|--------------|----------|------|------|------|------| | | | | | | | | | | | | | |
  
```

 Output
 ------

 The script outputs an Excel file containing the analyzed data. The path to the output file is specified in `output_file_path`. The format of the Excel file is:

 ```
   
 | 수요기관지역명 | 장비금액 | 계약금액 | 냉방용량 | 난방용량 | 날짜 |
 | --- | --- | --- | --- | --- | --- |
 |  |  |  |  |  |  |
   
 ```

 How it works
 ------------

 1. Reads the CSV file and selects the relevant columns for analysis.
 2. Changes the data type of certain columns to integers.
 3. Groups the data by '계약(납품요구)번호'.
 4. Filters the groups to find those that contain '냉난방기' in the '세부품명' column.
 5. For each filtered group, finds the kW information in the '품목' column, calculates the total kW information and device cost.
 6. Calculates the total construction cost for each group.
 7. Extracts the location and date information for each group.
 8. Stores the extracted information for each group in a new DataFrame.
 9. Outputs the new DataFrame as an Excel file.

 How to use
 ----------

 1. Install the required dependencies.
 2. Save the procurement data for a specific item in a CSV file and specify its path in `input_file_path`.
 3. Specify the path for the output file in `output_file_path`.
 4. Run the script. The analyzed data will be outputted to an Excel file at the specified path.
