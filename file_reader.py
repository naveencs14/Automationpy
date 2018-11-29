import csv
def read_csv_file(path):
    result = []
    try:
        with open(path, 'rb') as f:
            reader = csv.DictReader(f)
            result = list(reader)
            f.close()
    except Exception as e:
        print e
    finally:
        return result


