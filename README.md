# StyleFrame
_Exporting DataFrame to designed excel file have never been so easy_


A library that wraps pandas and openpyxl and allows easy styling of dataframes in excel
```
$ pip install StyleFrame
```
You can read the documentation at http://styleframe.readthedocs.org/en/latest/


## Usage Examples

First, lets create DataFrame that contains data we would like to export to .xlsx file 
```
import pandas as pd

df = pd.DataFrame({'Col A': [1, 20, -3],
                   'Col B': [15, 3, 116],
                   'Col C': [33, -6, 9]})

row_max_value = df.idxmax(axis=1)

df['Sum'] = df.sum(axis=1)
df['Mean'] = df.mean(axis=1)
```

Now, once we have the DataFrame ready, lets create a StyleFrame object
```
from StyleFrame.style_frame import StyleFrame

sf = StyleFrame(df)
```

The StyleFrame object will auto-adjust the columns width and the rows height
but we could change them manually
```
sf.set_column_width_dict(col_width_dict={
    ('Col A', 'Col B', 'Col C'): 15.3,
    ('Sum', 'Mean'): 30
})

'''
 excel rows starts from 1
 row number 1 is the headers
'''
sf.set_row_height_dict(row_height_dict={
    (1,): 45,
    (2, 3, 4): 25
})
```

Next, lets change the background colour of the maximum values to red and the font to white
```
from StyleFrame.styler import Styler, colors

for row_index, col_name in rows_max_value.iteritems():
    sf[col_name][row_index].style = Styler(bg_color=colors.red, bold=True, font_color=colors.white).create_style()
```

And change the font and the size of Sum and Mean columns
```
sf.apply_column_style(cols_to_style=['Sum', 'Mean'],
                      font_color='#40B5BF',
                      font_size=18,
                      bold=True,
                      style_header=True)
```

Finally, lets export to excel but not before we use the best features of StyleFrame:
- Change the page writing side
- Freeze rows and columns
- Add filters to headers

```
sf.to_excel('data.xlsx',
            right_to_left=False,
            columns_and_rows_to_freeze='B2', # will freeze the rows above 2 (=row 1 only) and columns that before column 'B' (=col A only)
            row_to_add_filters=0).save()
```

**_the result:_**
<img src="https://s32.postimg.org/gh2d7ruet/image.png">