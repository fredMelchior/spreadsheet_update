import gspread


def google_sheets_auth():
    """
    gspread Authentication with json file
    getting spreadsheet data by URL
    :return: spreadsheet as sh
    """
    gc = gspread.service_account(
        filename="/Users/fred/PycharmProjects/tunts-test-project/norse-voice-400513"
                 "-d0840a6c0d47"
                 ".json"
    )
    sh = gc.open_by_url(
        "SPREADSHEET_LINK_GOES_HERE"
    )

    return sh


def spreadsheet_data(sheet):
    """
    :param sheet:
    :return: spreadsheet data without the headers
    """
    return sheet.get_all_values()[3:]


def average_calculator(line_data, columns):
    """
    Calculate the grade point average of each student
    param: line_data, columns
    return: grade point average
    """

    grades = []

    for test in ["P1", "P2", "P3"]:
        try:
            grades.append(float(line_data[columns[test]]))
        except (KeyError, IndexError, ValueError):
            pass

    if grades:
        return sum(grades) / len(grades)
    else:
        return None


def situation_and_naf(average, absences, total_classes):
    """
    Function to set the student situation and points to be approved with a final test
    :param average:
    :param absences:
    :param total_classes:
    :return: situation and NAF
    """
    if 50 <= average < 70:
        return "Exame Final", round(max(0, 150 - average), 1)
    # elif 50 <= average < 70:
    #     return "Exame Final", max(0, round((14 - average) * 2))
    elif average < 50:
        return "Reprovado por Nota", 0
    elif absences > 0.25 * total_classes:
        return "Reprovado por Falta", 0
    else:
        return "Aprovado", 0


def update_spreadsheet(sheet, data, columns):
    """
    Final function that uses all information to update the spreadsheet with the new data
    :param sheet:
    :param data:
    :param columns:
    :return: NAF and Situation columns updated
    """
    for i, line_data in enumerate(data, start=4):
        average = round(average_calculator(line_data, columns), 1)
        absences = int(line_data[columns["Faltas"]])
        situation, naf = situation_and_naf(average, absences, total_classes=60)

        sheet.update_acell(f"G{i}", situation)
        sheet.update_acell(f"H{i}", str(naf))


def main():
    spreadsheet = google_sheets_auth()
    sheet = spreadsheet.get_worksheet(0)

    columns = {"Matricula": 0, "Nome": 1, "Faltas": 2, "P1": 3, "P2": 4, "P3": 5}

    data = spreadsheet_data(sheet)

    update_spreadsheet(sheet, data, columns)


if __name__ == "__main__":
    main()
