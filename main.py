import pandas as pd


def read_excel():
    df_source = pd.read_excel('source.xlsx', sheet_name='Sheet1')
    df_target = pd.read_excel('target.xlsx', sheet_name='Sheet1')
    # Sort the columns in the target dataframe to match the order in the source dataframe
    df_target = df_target.reindex(columns=df_source.columns)
    return df_source, df_target


def count_of_fields():
    df_source, df_target = read_excel()
    # create a new column in both dataframes combining the file name and extension
    df_source["FileName_Extension"] = df_source["FileName"]
    df_target["FileName_Extension"] = df_target["FileName"]

    # count occurrences of each file name and extension in the source and target data
    source_counts = df_source.groupby(["FileName_Extension"]).size().reset_index(name='Source_Count')
    target_counts = df_target.groupby(["FileName_Extension"]).size().reset_index(name='Target_Count')

    # merge the two counts on the file name and extension
    file_counts = pd.merge(source_counts, target_counts, how='outer', on='FileName_Extension')

    # fill NaN values with 0
    file_counts = file_counts.fillna(0)

    # add a new column named Outcome and fill it with values based on the comparison of Source_Count and Target_Count
    file_counts["Outcome"] = file_counts.apply(
        lambda row: "Matched" if row["Source_Count"] == row["Target_Count"] else "Not matched", axis=1)

    return file_counts


def compare_fields():
    df_source, df_target = read_excel()
    file_counts = count_of_fields()

    # Initialize dataframes to store the data that doesn't match or isn't found
    df_not_found = pd.DataFrame(columns=df_source.columns)
    df_mismatched = pd.DataFrame(columns=df_source.columns)
    df_identical = pd.DataFrame(columns=df_source.columns)
    df_target_mismatch = pd.DataFrame(columns=df_target.columns)
    df_source_not_pdf = pd.DataFrame()
    df_target_not_pdf = pd.DataFrame()

    # Check for non-PDF file extensions in the source dataframe
    for source_index, source_row in df_source.iterrows():
        file_extension = source_row['File_Extension']
        source_key = source_row['Key']
        if file_extension != '.pdf':
            df_source_not_pdf = df_source_not_pdf.append(source_row, ignore_index=True)

    # Check for non-PDF file extensions in the target dataframe
    for target_index, target_row in df_target.iterrows():
        file_extension = target_row['File_Extension']
        target_key = target_row['Key']
        if file_extension != '.pdf':
            df_target_not_pdf = df_target_not_pdf.append(target_row, ignore_index=True)

    # Check for matching key values in the source and target dataframes for PDF file extensions
    for source_index, source_row in df_source.iterrows():
        file_extension = source_row['File_Extension']
        source_key = source_row['Key']
        if file_extension == '.pdf':
            target_row = df_target.loc[df_target['Key'] == source_key]
            # If key is not found in target dataframe, add row to not found dataframe
            if target_row.empty:
                df_not_found = df_not_found.append(source_row, ignore_index=True)
            else:
                target_index = target_row.index[0]
                # If key is found in target dataframe, compare the rows
                if source_row.equals(df_target.iloc[target_index]):
                    df_identical = df_identical.append(source_row, ignore_index=True)
                else:
                    # If the rows don't match, add them to the mismatched dataframe
                    mismatched_row = source_row.copy()
                    for col in df_source.columns:
                        if not source_row[col] == df_target.iloc[target_index][col]:
                            mismatched_row[col] = f"{source_row[col]} --> {df_target.iloc[target_index][col]}"
                    df_mismatched = df_mismatched.append(mismatched_row, ignore_index=True)

    # Check for key values in target dataframe that are not in source dataframe for PDF file extensions
    for target_index, target_row in df_target.iterrows():
        target_key = target_row['Key']
        source_row = df_source.loc[df_source['Key'] == target_key]

        # If key is not found in source dataframe, add row to target mismatch dataframe
        if source_row.empty:
            df_target_mismatch = df_target_mismatch.append(target_row, ignore_index=True)

    # Write the dataframes to separate sheets in the output excel file
    with pd.ExcelWriter('output_file.xlsx') as writer:
        df_not_found.to_excel(writer, sheet_name='NotFound', index=False)
        df_mismatched.to_excel(writer, sheet_name='DataMismatch', index=False)
        df_identical.to_excel(writer, sheet_name='IdenticalData', index=False)
        df_target_mismatch.to_excel(writer, sheet_name='TargetMismatch', index=False)
        df_source_not_pdf.to_excel(writer, sheet_name='sourceOtherext', index=False)
        df_target_not_pdf.to_excel(writer, sheet_name='targetOtherext', index=False)
        file_counts.to_excel(writer, sheet_name='FileCounts', index=False)


if __name__ == '__main__':
    compare_fields()
