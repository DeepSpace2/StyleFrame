# StyleFrame
_Exporting DataFrame to designed excel file has never been so easy_


A library that wraps pandas and openpyxl and allows easy styling of dataframes in excel
```
$ pip install styleframe
```
You can read the documentation at http://styleframe.readthedocs.org/en/latest/


## Usage Examples

First, let us create a DataFrame that contains data we would like to export to an .xlsx file 
```python
import pandas as pd


columns = ['Date', 'Col A', 'Col B', 'Col C', 'Percentage']
df = pd.DataFrame(data={'Date': [date(1995, 9, 5), date(1947, 11, 29), date(2000, 1, 15)],
                        'Col A': [1, 2004, -3],
                        'Col B': [15, 3, 116],
                        'Col C': [33, -6, 9],
                        'Percentage': [0.113, 0.504, 0.005]},
                  columns=columns)

only_values_df = df[columns[1:-1]]

rows_max_value = only_values_df.idxmax(axis=1)

df['Sum'] = only_values_df.sum(axis=1)
df['Mean'] = only_values_df.mean(axis=1)
```

Now, once we have the DataFrame ready, lets create a StyleFrame object
```python
from StyleFrame import StyleFrame

sf = StyleFrame(df)
# it is also possible to directly initiate StyleFrame
sf = StyleFrame({'Date': [date(1995, 9, 5), date(1947, 11, 29), date(2000, 1, 15)],
                 'Col A': [1, 2004, -3],
                 'Col B': [15, 3, 116],
                 'Col C': [33, -6, 9],
                 'Percentage': [0.113, 0.504, 0.005]})
```

The StyleFrame object will auto-adjust the columns width and the rows height
but they can be changed manually
```python
sf.set_column_width_dict(col_width_dict={
    ('Col A', 'Col B', 'Col C'): 15.3,
    ('Sum', 'Mean'): 30,
    ('Percentage', ): 12
})

# excel rows starts from 1
# row number 1 is the headers
# len of StyleFrame (same as DataFrame) does not count the headers row
all_rows = tuple(i for i in range(1, len(sf) + 2))
sf.set_row_height_dict(row_height_dict={
    all_rows[0]: 45,
    all_rows[1:]: 25
})
```

Applying number formats
```python
from StyleFrame import Styler, utils


sf.apply_column_style(cols_to_style='Date',
                      styler_obj=Styler(number_format=utils.number_formats.date, font='Calibri', bold=True))

sf.apply_column_style(cols_to_style='Percentage',
                      styler_obj=Styler(number_format=utils.number_formats.percent))

sf.apply_column_style(cols_to_style=['Col A', 'Col B', 'Col C'],
                      styler_obj=Styler(number_format=utils.number_formats.thousands_comma_sep))
                      
# if using version < 0.2 you need to define the style with style specifiers
sf.apply_column_style(cols_to_style=['Col A', 'Col B', 'Col C'],
                      number_format=utils.number_formats.thousands_comma_sep)

```

Next, let's change the background color of the maximum values to red and the font to white  
we will also protect those cells and prevent the ability to change their value
```python
style = Styler(bg_color=utils.colors.red, bold=True, font_color=utils.colors.white, protection=True,
               underline=utils.underline.double, number_format=utils.number_formats.thousands_comma_sep).create_style()
for row_index, col_name in rows_max_value.items():
    sf[col_name][row_index].style = style
```

And change the font and the font size of Sum and Mean columns
```python
sf.apply_column_style(cols_to_style=['Sum', 'Mean'],
                      styler_obj=Styler(font_color='#40B5BF',
                                        font_size=18,
                                        bold=True),
                      style_header=True)
# if using version < 0.2 you need to define the style with style specifiers
sf.apply_column_style(cols_to_style=['Sum', 'Mean'],
                      font_color='#40B5BF',
                      font_size=18,
                      bold=True,
                      style_header=True)

```

Change the background of all rows where the date is after 14/1/2000 to green
```python                 
sf.apply_style_by_indexes(indexes_to_style=sf[sf['Date'] > date(2000, 1, 14)],
                          cols_to_style='Date',
                          styler_obj=Styler(bg_color='green', number_format=utils.number_formats.date, bold=True))
# if using version < 0.2 you need to define the style with style specifiers
sf.apply_style_by_indexes(indexes_to_style=sf[sf['Date'] > date(2000, 1, 14)],
                          cols_to_style='Date',
                          bg_color=utils.colors.green,
                          number_format=utils.number_formats.date,
                          bold=True)
              
```

Finally, let's export to Excel but not before we use more of StyleFrame's features:
- Change the page writing side
- Freeze rows and columns
- Add filters to headers

```python
ew = StyleFrame.ExcelWriter('sf tutorial.xlsx')
sf.to_excel(excel_writer=ew,
            sheet_name='1',
            right_to_left=False,
            columns_and_rows_to_freeze='B2', # will freeze the rows above 2 (=row 1 only) and columns that before column 'B' (=col A only)
            row_to_add_filters=0,
            allow_protection=True)
```

Adding another excel sheet
```python
other_sheet_sf = StyleFrame({'Dates': [date(2016, 10, 20), date(2016, 10, 21), date(2016, 10, 22)]},
                            styler_obj=Styler(number_format=utils.number_formats.date))
# if using version < 0.2 you need to define the style with style specifiers
other_sheet_sf = StyleFrame({'Dates': [date(2016, 10, 20), date(2016, 10, 21), date(2016, 10, 22)]},
                            number_format=utils.number_formats.date)

other_sheet_sf.to_excel(excel_writer=ew, sheet_name='2')
```

Don't forget to save
```python
ew.save()
```

**_the result:_**
<img src="https://s10.postimg.org/ppt8gt5m1/Untitled.png">
