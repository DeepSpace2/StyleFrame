import pandas as pd
from datetime import timedelta
from StyleFrame import StyleFrame, Styler, utils

# data.csv contains the first 500 rows of Kaggle's "StackLite: Stack Overflow questions and tags"
# dataset available at https://www.kaggle.com/stackoverflow/stacklite
df = pd.read_csv('data.csv', parse_dates=['CreationDate', 'ClosedDate', 'DeletionDate'])

sf = StyleFrame(df)

# Using red background for Id column for rows with questions that were closed less than 5 minutes after creation
sf.apply_style_by_indexes(indexes_to_style=sf[sf['ClosedDate'] - sf['CreationDate'] < timedelta(minutes=5)],
                          styler_obj=Styler(bg_color=utils.colors.red),
                          cols_to_style=['Id'])

# Changing the width of the date columns so their content fits nicely
sf.set_column_width(columns=['CreationDate', 'ClosedDate', 'DeletionDate'],
                    width=20)

# Using color-scale conditional formatting for the questions' scores, based on percentage
sf.add_color_scale_conditional_formatting(start_type=utils.conditional_formatting_types.percentile,
                                          start_value=0,
                                          start_color=utils.colors.red,
                                          end_type=utils.conditional_formatting_types.percentile,
                                          end_value=100,
                                          end_color=utils.colors.green,
                                          columns_range=['Score'])


# Adding filters to the header row, freezing it and exporting to Excel
sf.to_excel('output.xlsx', columns_and_rows_to_freeze='A2', row_to_add_filters=0,
            best_fit=['OwnerUserId', 'AnswerCount']).save()
