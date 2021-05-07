# Microsoft Teams Attendance

This python script made for calculate the attendance date (duration, joining count, join timestamp, leave timestamp) of student.

This script is use `date  modified` as end class time.

![date modified](images/date_modified_sample.png)

## Usage
- Install pandas package.
```bash
pip install pandas
```
- Run script file.
```bash
python script.py "<csv_filename>"
```
the name of output file is `Output_<csv_filename>`.
 
## Example output file
![output example](images/output_sample.png)