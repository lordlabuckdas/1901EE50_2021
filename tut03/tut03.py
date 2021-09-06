import os

# CSV data file
REGISTRATION_CSV_FILE = "regtable_old.csv"


def create_folder(folder_name):
    """
    Create folder if it does not exist already
    """
    if not os.path.exists(folder_name):
        os.makedirs(folder_name)


def parse_csv(raw_data):
    """
    Parse CSV data into list of lists
    """
    raw_data = raw_data.split("\n")
    data = []
    for row in raw_data:
        if row:
            data.append(row.split(","))
    return data


def write_data(folder, key, data):
    """
    Write csv data into folder by key (for file name)
    """
    # list for column names
    columns = []
    for column in data[0]:
        columns.append(column)
    # key_idx denotes column index in CSV data
    key_idx = 0
    if key == "roll_num":
        key_idx = 0
    elif key == "subno":
        key_idx = 3
    else:
        raise Exception("Invalid key!")
    del data[0]
    # parse data row-wise
    for row in data:
        key_val = row[key_idx]
        # if key_val is none, continue to next iteration
        if not key_val:
            continue
        file_name = os.path.join(folder, key_val + ".csv")
        output_file = open(file_name, "a")
        # if empty file, insert column names
        if not os.stat(file_name).st_size:
            output_file.write(
                ",".join([columns[0], columns[1], columns[3], columns[8]])
            )
            output_file.write("\n")
        # insert data as a row
        output_file.write(",".join([row[0], row[1], row[3], row[8]]))
        output_file.write("\n")
        output_file.close()


def output_by_subject():
    """
    Subject-wise output csv files
    """
    create_folder("output_by_subject")
    reg_csv_file = open(REGISTRATION_CSV_FILE, "r")
    data = parse_csv(reg_csv_file.read())
    reg_csv_file.close()
    write_data("output_by_subject", "subno", data)


def output_individual_roll():
    """
    Roll number-wise output csv files
    """
    create_folder("output_individual_roll")
    reg_csv_file = open(REGISTRATION_CSV_FILE, "r")
    data = parse_csv(reg_csv_file.read())
    reg_csv_file.close()
    write_data("output_individual_roll", "roll_num", data)


output_individual_roll()
output_by_subject()
