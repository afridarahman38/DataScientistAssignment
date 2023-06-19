# DataScientistAssignment
### Data Scientist Assignment Solution and Explanation of the Process <br>

#### Background of the Assignment
There are excel file, where 2 sheets are attached. In sheet 1 there are information about:
- ID	
- Customer Name	
- Division	
- Gender	
- MaritalStatus 
- Age
- Income

and in sheet 2 there are information about 
- ID


### Assignment Part 1: Excel

In assignment, it is said to

3. Create a pivot table using data of sheet1 and show the information following
this structure -
- a. The data table should show the Sum of Income as value.
- b. The columns should include the value of Gender and MaritalStatus.
- c. The rows should be in the following order: Division; Customer Name; ID.

4. In sheet2 there are some IDs. Add a new column to sheet1 and name it
“Matched”. Please match the IDs of sheet2 with the IDs of sheet1 and show
the result as True or False. You need to use a formula to match the IDs.



### Assignment Part 2: Python

1. In this part, separate the sheet 1 data in a new excel sheet.
2. Convertion the excel file to csv.
3. Load the csv file.
4. Get rid of the column ID from the data frame.
5. Encode the data to have similar values.
6. Now use K-means clustering based on their divisions. This part is a bonus
task.
7. Download the file without omitting the output.

3. Load the CSV file.
4. Get rid of the column ID from the data frame.
5. Encode the data to have similar values.
6. Now use K-means clustering based on their divisions. This part is a bonus
task.
7. Download the file without omitting the output.


## Installation
Used colab to do Part 2 task.

Installed the version of Python and the environment setup I am using:

```bash
    Python 3.10.9
    anaconda 1.11.2
    pip 22.3.1
    xlwings 0.29.1 
```

## Explanation of Process Part 1:
#### To fix "Age" column RANDBETWEEN Excel formula
#### formula items:

| Parameter | Type     | Description                |
| :-------- | :------- | :------------------------- |
| `Age` | `=RANDBETWEEN` | **generate random numbers between two |

#### To fix this item

| Parameter | Type     | Description                       |
| :-------- | :------- | :-------------------------------- |
| `Age`      | `value` | **Required**. Age is converted to value |

### 3. Solution

- a. Used pivot table for "Sum of Income as value"
- b. Used pivot table for the "value of Gender" and "Value of MaritalStatus"
- c. "Rows should be in the following order: Division; Customer Name;
ID" - In this problem faced a problem with using the Customer Name column.

## Explanation of Process Part 2:
### 1. Separate the sheet 1 data in a new Excel sheet.

For this part I used "PyCharm". Tried to do on colab, but faced the problem. For this part I had to install anaconda, xlwings. 

`xlwings`

`anaconda`

As xlwings is a open-core spreadsheet automation package with a beautiful API. It made the code run correctly. 
After run the code from "Command Prompt"

`cd filename`

`python filename.py`

### Run Locally

Clone the project

```bash
  git clone [https://github.com/afridarahman38/DataScientistAssignment/tree/master]
```

Code of Part 2: 1

```bash
  import xlwings as xw

EXCEL_FILE = 'AssignmentforDataScientist.xlsx'

try:
    excel_app = xw.App(visible=False)
    wb = excel_app.books.open(EXCEL_FILE)
    for sheet in wb.sheets:
        sheet.api.Copy()
        wb_new = xw.books.active
        wb_new.save(f'{sheet.name}.xlsx')
        wb_new.close()

finally:
    excel_app.quite()
```

In Command Prompt

```bash
  cd filename
  python filename.py
```

Output

```bash
  Sheet1.xlsx
  Sheet1.xlsx
```
### 2. Convert the excel file to csv.

```bash
    import pandas as pd

    df = pd.read_excel('/content/Sheet1.xlsx', sheet_name='Sheet1') #read excel file

    df.to_csv('/content/Sheet1.csv', index=False) #excel to csv
```

Output

```bash
    Sheet1.xlsx
    Sheet1.csv
```
### 3. Now load the csv file.

```bash
    import pandas as pd

    df = pd.read_csv('/content/Sheet1.csv') # load csv file
```

Output

```bash
    It loaded the Sheet1.csv file.
```
### 4. Encode the data to have similar values.

```bash
    df = df.drop('ID', axis=1) #get rid of the column ID

print(df)
```

Output

```bash
    encoded the dataset.
```
### 5. Encode the data to have similar values.

```bash
     from sklearn.preprocessing import LabelEncoder

    encoder = LabelEncoder() #creating instance in LabelEncoder

    df['Division'] = encoder.fit_transform(df['Division']) #encode the column

    print(df)
```

Output

```bash
    encodeed the dataset by Division
```
### 6. Now use K-means clustering based on their divisions. This part is a bonus
task.

```bash
    import pandas as pd
from sklearn.cluster import KMeans

df = pd.read_csv('/content/Sheet1.csv') #load dataset

print(df['Division'].unique()) #verify unique values

features = ['Division'] #selecting column

kmeans = KMeans(n_clusters=8)  # number of clusters
kmeans.fit(df[features])

cluster_labels = kmeans.labels #retrieve cluster labels to data point

df['Cluster_Labels'] = cluster_labels #add cluster labels to dataframe
```

Output

```bash
        faced problem
```

### 7. Download the file without omitting the output.

```bash
    files.download('/content/Sheet1.csv') #for download
```

Output

```bash
    Sheet1.csv downloaded
```

## Contact

Your Name - Afrida Rahman - https://www.linkedin.com/in/afrida-rahman-152287199/ - afridaurmi@gmail.com

Project Link: https://github.com/afridarahman38/DataScientistAssignment/tree/master







