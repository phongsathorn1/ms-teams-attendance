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
import sys

if __name__ == "__main__":
    filename = sys.argv[1]
    
    # Get modified time of file (Teacher download file when end class)
    fname = pathlib.Path(filename)
    end_datetime = datetime.datetime.fromtimestamp(fname.stat().st_mtime)

    # Attendance file download from MS Teams using UTF-16 encoding
    df = pd.read_csv(filename, encoding='UTF-16', sep='\t')
    df['Timestamp'] = df['Timestamp'].astype('datetime64[ns]')

    result_dict = {
        "Full Name": [],
        "Duration": [],
        "Join Counts": [],
        "From Timestamp": [],
        "To Timestamp": []
    }

    names = df['Full Name'].unique()

    for name in names:
        user_df = df[df["Full Name"] == name]
        
        last_start_time = None
        sum_time = None
        is_join = False
        join_count = 0
        
        for index, row in user_df.iterrows():
            action = row['User Action']
            timestamp = row['Timestamp']

            if action == 'Joined':
                last_start_time = timestamp
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

        result_dict["Full Name"].append(user_df['Full Name'].iloc[0])
        result_dict["Duration"].append(sum_time)
        result_dict["Join Counts"].append(join_count)
        result_dict["From Timestamp"].append(from_timestamp)
        result_dict["To Timestamp"].append(to_timestamp)

    result_df = pd.DataFrame(result_dict)
    result_df.to_csv("Output_%s" %(filename), index=False)
