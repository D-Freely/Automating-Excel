# Automating-Excel

Python script which automates a time-consuming manual excel task

Python libraries used:

- Openpyxl
- Pandas
- Matplotlib
- Numpy

See the _code.py_ file for the python code. 

### Description

This python script takes a full table of raw financial data and filters out and saves a separate excel document for each work-stream that is mentioned in the original file (including the relevant excel formulas). It does this while cleaning the formatting of the original excel file. It also adds in a 'Total Variance' column and creates a bar chart visualisation of forecast vs actuals in a second sheet per workstream filtered excel file. The script has been coded so that it will be able to accept new columns added into the original file for future months.

#### The first few rows of the original file look this:

![image](https://user-images.githubusercontent.com/92688098/141981248-865a9798-576e-44ad-8a61-c1b60be98fec.png)

#### An example of one of the output files looks like this:

Sheet 1:
![image](https://user-images.githubusercontent.com/92688098/141981135-1611fd9d-6f4f-48f0-b1f7-ed28b011074b.png)

Sheet 2:
![image](https://user-images.githubusercontent.com/92688098/141980945-3426d28c-a5b2-470d-b446-f4a91a7932fb.png)
