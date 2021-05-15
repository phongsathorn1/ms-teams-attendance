"""
MIT License

Copyright (c) 2021 Phongsathorn Kittiworapanya

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.
"""

import pandas as pd
import datetime
import pathlib
from glob import glob
import os

if __name__ == "__main__":
    studentlist_file = input()
    professsor_name = input().lower()

    writer = pd.ExcelWriter('attendance.xlsx')
    
    for filename in glob("csv_files/*.csv"):
        # Get modified time of file (Teacher download file when end class)
        fname = pathlib.Path(filename)
        end_datetime = datetime.datetime.fromtimestamp(fname.stat().st_mtime)

        # Attendance file download from MS Teams using UTF-16 encoding
        df = pd.read_csv(filename, encoding='UTF-16', sep='\t')
        df['Timestamp'] = df['Timestamp'].astype('datetime64[ns]')
        start_datetime = df[df['Full Name'].str.lower() == professsor_name].iloc[0]["Timestamp"]

        result_dict = {
            "Full Name": [],
            "Duration": [],
            "Join Counts": [],
            "From Timestamp": [],
            "To Timestamp": []
        }

        names = df['Full Name'].unique()

        for name in names:
            user_df = df[df["Full Name"] == name].reset_index(drop=True)
            
            last_start_time = None
            sum_time = None
            is_join = False
            join_count = 0
            
            for index, row in user_df.iterrows():
                action = row['User Action']
                timestamp = row['Timestamp']

                if action in ['Joined', 'Joined before']:
                    last_start_time = timestamp
                    if last_start_time < start_datetime and index == 0:
                        last_start_time = start_datetime
                    is_join = True
                    join_count += 1

                elif action == 'Left':
                    is_join = False
                    if not isinstance(sum_time, datetime.timedelta):
                        sum_time = row['Timestamp'] - last_start_time
                    else:
                        sum_time += row['Timestamp'] - last_start_time

            if is_join:
                if not isinstance(sum_time, datetime.timedelta):
                    sum_time = end_datetime - last_start_time
                else:
                    sum_time += end_datetime - last_start_time
                to_timestamp = end_datetime
            else:
                to_timestamp = user_df['Timestamp'].iloc[-1]

            from_timestamp = user_df['Timestamp'].iloc[0]

            result_dict["Full Name"].append(' '.join(user_df['Full Name'].iloc[0].split()))
            result_dict["Duration"].append(int(sum_time.total_seconds() / 60))
            result_dict["Join Counts"].append(join_count)
            result_dict["From Timestamp"].append(from_timestamp)
            result_dict["To Timestamp"].append(to_timestamp)

        result_df = pd.DataFrame(result_dict)

        class_df = pd.read_excel(studentlist_file, sheet_name='CP63')

        eng_names = []
        for en_name in class_df['eng name']:
            eng_names.append(' '.join(en_name.replace('Miss', '').replace('Mr.', '').replace('Ms.', '').split()))
        class_df['eng name'] = eng_names

        result_df = result_df.rename(columns={"Full Name": "eng name"})
        result_df = pd.merge(class_df, result_df, how='left', on='eng name')
        professsor_df = pd.DataFrame({
            'student_id': [None],
            'thai name': [None],
            'eng name': [professsor_name],
            'Join Counts': [None],
            'Duration': [int((end_datetime - start_datetime).total_seconds() / 60)],
            'From Timestamp': [start_datetime], 
            'To Timestamp': [end_datetime]
            })

        result_df = pd.concat([professsor_df, result_df], ignore_index=True)
        result_df.to_excel(writer, sheet_name="%s" %(os.path.basename(filename).replace(".csv", "").replace("-meetingAttendanceList", "-attendance")), index=False, encoding="utf-8")

    writer.save()
