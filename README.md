# DataFill
A classification model training framework with high scalability.
## Environment
```
python=3.6.13
openpyxl==3.1.2
pillow==8.4.0
pyyaml==6.0.1
pyinstaller==4.10 (Used to package as .exe file)
```
## Instructions
### 1. Create data file
You can refer to the data.xlsx example. In the file, each line of data is used to fill in the template once. The first line is the name of the data, which is required but not entered in the template. The data in the first column is used as the file name of the output result, so the value must be unique.
### 2. Create template file
Custom Excel files based on your data needs.
### 3. Create config file
The file is in yaml format. In this file, you can set which data in the data file to copy to the template file in which sheets in which positions, you can also set the font style, size, color, time display format and so on. Refer to the comments in the config.yaml file for details.
### 4. run
```bash
python start.py
```
When the program is running, enter the absolute path of the corresponding file according to the prompt.
## Demonstrate
```bash
python start.py
```
![运行界面](https://github.com/junnnier/DataFill/blob/main/run_image.png)
## Pack
You can package the program as DataFill.exe file for use on windows systems.
```bash
pyinstaller -D -n "DataFill" start.py
```