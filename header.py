import pandas as pd
import numpy as np
import re
import sys

filepath = sys.argv[3]
df = pd.read_excel(filepath, header=None)
#根据表格内容设计定要求合并的表头的行数和列数
header_row_num = int(sys.argv[1])
header_col_num = int(sys.argv[2])
#合并水平表头
def combine_header(df, header_row_num, header_col_num):
    def combine_list(data, axis):
        def remove_duplicates(lst):
            seen = set()
            result = []
            for item in lst:
                if item not in seen:
                    seen.add(item)
                    result.append(item)
            return result
        
        results = []
        if axis == 0:
            array = data.to_numpy().T
        else:
            array = data.to_numpy()
        for l in array:
            results.append(','.join(remove_duplicates(list(map(lambda x:str(x), l)))))
        return results
    df_h = df[0:header_row_num]
    df.iloc[header_row_num-1, :] = combine_list(df_h, 0)
    df = df.iloc[header_row_num-1:,:]
    df_v = df.iloc[:,:header_col_num]
    df.iloc[:,header_col_num-1] = combine_list(df_v, 1)
    df = df.iloc[:, header_col_num-1:]
    #print(df[:,header_col_num])
    return df

df = combine_header(df, header_row_num, header_col_num)
df = df.map(lambda x:str(x).replace('\n', ''))
#df.to_string(filepath.replace('.xlsx', '.txt'), index=False)
df = df.to_string(index=False)
df = re.sub(r'[^\S\n]+', ' ', df)
df = re.sub(r'\n\s*', '\n', df)
with open(filepath.replace('.xlsx', '.txt'), 'w', encoding='utf8') as f:
    f.write(df)