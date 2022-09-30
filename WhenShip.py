from itertools import count
import os
from datetime import date
from os.path import exists
from sys import path
from file_read_backwards import FileReadBackwards
import gspread
import webbrowser
from gspread.cell import Cell
import emoji


credentials = {"installed": {"client_id": "703535930585-4vg06hqfeuc3tb2cq5s922roag77tv5u.apps.googleusercontent.com", "project_id": "steel-signifier-354722", "auth_uri": "https://accounts.google.com/o/oauth2/auth",
                             "token_uri": "https://oauth2.googleapis.com/token", "auth_provider_x509_cert_url": "https://www.googleapis.com/oauth2/v1/certs", "client_secret": "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", "redirect_uris": ["http://localhost"]}}
authorized_user = {"refresh_token": "1//06dXh3D-LNMm2CgYIARAAGAYSNwF-L9IrQDkhaKvfE8fhpCWSKngPohcsHgtcFgi20goY7YZtRRYkbYRUkKgOzN21b_SIv4zoaQM", "token_uri": "https://oauth2.googleapis.com/token", "client_id": "703535930585-4vg06hqfeuc3tb2cq5s922roag77tv5u.apps.googleusercontent.com",
                   "client_secret": "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", "scopes": ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"], "expiry": "2022-08-01T23:28:47.235664Z"}

gc, authorized_user = gspread.oauth_from_dict(credentials, authorized_user)

URL = 'https://docs.google.com/spreadsheets/d/1r47wk5xybRIdfjW53aPgKQWvabszQ8QEGycy3WxxIRc/edit#gid=0'
sh = gc.open_by_url(URL)
cameras = []

CURRENT_YEAR = date.today().year
PRODUCTION_CATEGORIES = [r"CCD\004", r"CCD\009", r"CCD\012", "CMOS"]
QUALITY_CATEGORIES = ["CCD", "CMOS"]

TESTS_RANKS = {
    "Firmware Start : OK": 0,
    "Saturation : OK": 1,
    "Soak1 : OK": 2,
    "Encoder : OK": 3,
    "Dark : OK ": 4,
    "Bright1 : OK": 5,
    "Object : OK": 6,
    "Bright2 : OK": 7,
    "LightLeak_Vibration : OK": 8,
    "Soak2 : OK": 9,
    "Firmware End : OK": 10,
    "Record PLEORA : OK": 11,
    "ABOUT THIS BUILD": 12
}


tests_col_locations = {
    "Not Started": 2,
    "Firmware Start": 3,
    "Saturation": 4,
    "Soak1": 5,
    "Encoder": 6,
    "Dark": 7,
    "Bright1": 8,
    "Object": 9,
    "Bright2": 10,
    "LightLeak_Vibration": 11,
    "Soak2": 12,
    "Firmware End": 13,
    "Record PLEORA": 14,
    "ABOUT THIS BUILD": 15,
    "Test Folder": 16
}


# Python code to merge dict using update() method
def merge(dict1, dict2):
    return(dict2.update(dict1))


def find_the_latest_test(filename):
    latest_test_rank = 0
    with open(filename) as f:
        contents = f.read()
        for test in TESTS_RANKS:
            if test in contents and TESTS_RANKS[test] > latest_test_rank:
                latest_test_rank = TESTS_RANKS[test]
    return list(TESTS_RANKS.keys())[list(
        TESTS_RANKS.values()).index(latest_test_rank)].replace(" : OK", "")


def add_to_camera_list(path_to_testdata):

    if "production" in path_to_testdata.lower():
        folder_name = path_to_testdata.split('\\')[-2]
    else:
        folder_name = path_to_testdata.split('\\')[-3]

    serial_number = folder_name.split(" ")[-1]
    test_log_filename = serial_number + "_TestLog.txt"

    # IF no testlog.txt in TestData, camera not started
    if test_log_filename not in os . listdir(path_to_testdata):
        camera = {"name": folder_name, "lastest_test": 'Not Started',
                  "test_folder": path_to_testdata}
        cameras.append(camera)
    else:
        # find lastest test and path_to_testdata folder.
        camera = {"name": folder_name,
                  "lastest_test": find_the_latest_test(f'{path_to_testdata}\{test_log_filename}'),
                  "test_folder": path_to_testdata}
        cameras.append(camera)


def check_objects_analysis(path_to_object_scan):
    objects_to_be_analyzed = {
        "crosstalk": "xtalk_",
        "stepgauge": "manu_",
        "grill": "algn_",
        "mtf": "",
    }

    checked_list = {}

    for file in os.listdir(path_to_object_scan):
        file_name = file.lower().strip().replace(" ", "")
        for object in objects_to_be_analyzed.keys():
            if object in file_name and "bar" not in file_name:
                if file_name.endswith((".tif", ".pgm")):
                    file_name = file_name.replace("tif", "png")
                    file_name = file_name.replace("pgm", "png")
                    if "mtf" not in file_name:
                        checked_list[objects_to_be_analyzed[object] +
                                     file_name] = "Missing"
                    else:
                        checked_list["mtf_horedge_plot.png"] = "Missing"
                        checked_list["mtf_horedge_result.png"] = "Missing"
                        checked_list["mtf_veredge_plot.png"] = "Missing"
                        checked_list["mtf_veredge_result.png"] = "Missing"
                elif file_name.endswith(".png"):
                    if file_name in checked_list:
                        checked_list.pop(file_name)
    return checked_list


def lightleak_vibration_analysis(path_to_lightleak_vibration):
    checked_list = {}
    for file in os.listdir(path_to_lightleak_vibration):
        file_name = file.lower().strip().replace(" ", "")
        if file_name.endswith((".tif", ".pgm")):
            file_name = file_name.replace("tif", "png")
            file_name = file_name.replace("pgm", "png")
            checked_list[file_name] = "Missing"
        elif file_name.endswith(".png") and file_name in checked_list:
            checked_list.pop(file_name)
    return checked_list


def fullframe_analysis(path_to_bright_test):
    count_fullframes_png = 0
    count_fullframes_tif = 0

    for file in os.listdir(path_to_bright_test):
        file_name = file.lower().strip().replace(" ", "")
        if file_name.endswith(".png") and "fullframe" in file_name:
            count_fullframes_png += 1
        elif file_name.endswith(".tif") and "fullframe" in file_name:
            count_fullframes_tif += 1

    return count_fullframes_png == count_fullframes_tif


def count_words(path):
    # creating variable to store the
    # number of words
    number_of_words = 0
    with open(path, 'r') as file:

        # Reading the content of the file
        # using the read() function and storing
        # them in a new variable
        data = file.read()

        # Splitting the data into separate lines
        # using the split() function
        lines = data.split()
        # Adding the length of the
        # lines in our number_of_words
        # variable
        number_of_words += len(lines)

    # Printing total number of words
    return number_of_words


def file_is_modified(path_to_text_file, default_words):

    if os.path.exists(path_to_text_file):

        number_of_words = count_words(path_to_text_file)
        file_creation_time = os.path.getctime(path_to_text_file)
        file_modification_time = os.path.getmtime(path_to_text_file)

        return file_creation_time != file_modification_time and number_of_words > default_words

    return False


def internal_images_are_uploaded(path_to_internal_photo):
    # Return true there are photo in this directory
    return len(os.listdir(path_to_internal_photo)) > 1


def get_config_name(string):
    counter = 0
    config = ""
    starting = False
    for char in string:
        if char == '"':
            counter += 1
            starting = True
            continue
        if counter == 2:
            break
        if starting:
            config += char

    return config


def check_test_results(path_to_test_result):
    mismatched_config = {}
    pass


def check_for_missing_data(test_folder):

    print(serial_number)
    missing_data = {}
    try:
        WORDS_IN_PCB_BY_DEFAULT = count_words(
            r"\\xscan\x\Template\FolderTemplate\Outgoing\TestData\PCB Serial Number.txt")
    except:
        WORDS_IN_PCB_BY_DEFAULT = 47
    WORDS_IN_XRAY_LOG_BY_DEFAULT = 51

    # Check Objects
    missing_analyzed_objects = check_objects_analysis(
        test_folder + r"\OBJECT")
    merge(missing_analyzed_objects, missing_data)

    # Check LightLeak_Vibration
    missing_lightleak_vibration = lightleak_vibration_analysis(
        test_folder + r"\LIGHTLEAK_VIBRATION")
    merge(missing_lightleak_vibration, missing_data)

    # Check FullFrame (CCD only)
    if "CCD" in test_folder:
        if not fullframe_analysis(test_folder + r"\BRIGHT1"):
            merge({"BRIGHT 1 Fullframe": "MISSING"},
                  missing_data)
        if not fullframe_analysis(test_folder + r"\BRIGHT2"):
            merge({"BRIGHT 2 Fullframe": "MISSING"},
                  missing_data)

    # Check PCB_Serial_Number
    if not file_is_modified(test_folder + r"\PCB Serial Number.txt", WORDS_IN_PCB_BY_DEFAULT):
        merge({"PCB_Serial_number.txt": "IS NOT FILLED YET"},
              missing_data)

    # Check Xray-Log
    if not file_is_modified(test_folder + r"\OBJECT\X-rayLog.txt", WORDS_IN_XRAY_LOG_BY_DEFAULT):
        merge({"X-rayLog.txt": "IS NOT FILLED YET"},
              missing_data)

    # Check TestResult.txt
    missing_results = check_test_results()

    # Check for internal pictures
    path_to_internal_photo = test_folder.replace(r"\TestData", r"\Photo")
    for folder in os.listdir(path_to_internal_photo):
        if "internal" in folder.lower():
            path_to_internal_photo += rf"\{folder}"
            break
    if not internal_images_are_uploaded(path_to_internal_photo):
        merge({"INTERNAL PHOTO": "NOT UPLOADED"}, missing_data)

    # merge all missing data analysis to a parent dictionary

    return missing_data


def is_proper_name(folder_name):
    try:
        so = folder_name.split(" ")[0].split("-")[0]
        part_number = folder_name.split(" ")[1]
    except:
        return False

    if not so.isdigit() or not part_number.startswith("X"):
        return False

    return True


def update_progress(in_progress_cameras, camera_location, worksheet, sale_category):
    for camera in in_progress_cameras:
        if sale_category == "production":
            path = f'{camera_location}\{camera}\TestData'
        else:
            oqc_folder = ""
            for folder in os.listdir(f'{camera_location}\{camera}'):
                if "incoming" not in folder and "failure" not in folder and not folder.endswith(".docx") and not folder.endswith(".pdf"):
                    oqc_folder = folder
            path = f'{camera_location}\{camera}\{oqc_folder}\TestData'
        if is_proper_name(camera):
            add_to_camera_list(path)

    # Add to Google Sheet
    current_row = 2
    cells = []
    DATA_ANALYSIS_COL = 17
    READY_TO_SHIP_COL = 18
    check_mark = ':check_mark:'
    cross_mark = ':cross_mark:'

    if len(cameras) != 0:
        for camera in cameras:
            lastest_test = camera["lastest_test"]
            current_col = tests_col_locations[lastest_test]
            test_folder_col = tests_col_locations["Test Folder"]
            # If testing is done, check for missing data analysis
            try:
                if lastest_test == "ABOUT THIS BUILD":
                    missing_data = check_for_missing_data(
                        camera["test_folder"])

                    if len(missing_data) == 0:
                        # Check for done
                        cells.append(Cell(row=current_row, col=DATA_ANALYSIS_COL,
                                          value=emoji.emojize(check_mark)))
                        cells.append(Cell(row=current_row, col=READY_TO_SHIP_COL,
                                          value=emoji.emojize(check_mark)))
                    else:
                        # Update Data Analysis Column
                        missing_items = ""
                        for item in missing_data:
                            missing_items += item + ": " + \
                                missing_data[item] + "\n"
                        cells.append(
                            Cell(row=current_row, col=DATA_ANALYSIS_COL, value=missing_items.upper()))
                        cells.append(Cell(row=current_row, col=READY_TO_SHIP_COL,
                                          value=emoji.emojize(cross_mark)))
            except:
                cells.append(Cell(row=current_row, col=1,
                             value="Failed to Generate Sheet"))
            # Update Camera Name in A1
            cells.append(Cell(row=current_row, col=1, value=camera["name"]))
            # Update the lastest test Associated with this camera
            cells.append(Cell(row=current_row, col=current_col,
                         value=(emoji.emojize(check_mark))))
            # Update Test Folder location
            cells.append(Cell(row=current_row, col=test_folder_col,
                         value=camera["test_folder"]))

            current_row += 1  # Go to the next row
        cells.append(Cell(row=current_row, col=1, value="End"))
        worksheet.update_cells(cells)


def get_inprogress_cams(camera_location):
    in_progress_cameras = []
    test_report_directory = ""

    for camera_name in os.listdir(camera_location):
        has_test_report = False
        path_to_camera_name = os.path.join(camera_location, camera_name)
        if os.path.isdir(path_to_camera_name):
            if is_proper_name(camera_name):
                if "quality" in path_to_camera_name.lower():
                    oqc_folder = ""
                    for folder in os.listdir(path_to_camera_name):
                        if "incoming" not in folder and "failure" not in folder and not folder.endswith(".docx") and not folder.endswith(".pdf"):
                            oqc_folder = folder

                    certificate_directory = path_to_camera_name + \
                        rf'\{oqc_folder}\Doc\USB\Detector File\Doc\Certificate of Conformance.pdf'
                    test_report_directory = path_to_camera_name + \
                        rf'\{oqc_folder}\TestData'

                else:
                    certificate_directory = path_to_camera_name + \
                        r"\Doc\USB\Detector File\Doc\Certificate of Conformance.pdf"
                    test_report_directory = path_to_camera_name + \
                        r"\TestData"

                # Check if CoC is is USB
                has_coc_pdf = exists(certificate_directory)

                # Check if Test Report is in TestData folder
                for file in os.listdir(test_report_directory):
                    if "test report.pdf" in file.lower():
                        has_test_report = True
                        break

                # Both Test Report and CoC must present if a camera is considered DONE
                if not has_coc_pdf and not has_test_report:
                    in_progress_cameras.append(camera_name)

    return in_progress_cameras


def main():
    print("Camera Building Progress Report")
    webbrowser.open(URL)
    production_root = rf"\\xscan\X\Production\{CURRENT_YEAR}\Detector"
    for category in PRODUCTION_CATEGORIES:
        camera_location = f'{production_root}\{category}'
        # Open the worksheet
        worksheet = sh.worksheet(category)
        worksheet.batch_clear(["A2:Z50"])  # Clear the Sheet from A2 to Z50
        in_progress_cameras = get_inprogress_cams(camera_location)
        update_progress(in_progress_cameras, camera_location,
                        worksheet, "production")
        cameras.clear()

    quality_root = rf"\\xscan\X\Quality\{CURRENT_YEAR}"
    for category in QUALITY_CATEGORIES:
        camera_location = f'{quality_root}\{category}'
        # Open the worksheet
        worksheet = sh.worksheet(category + " RMA")
        worksheet.batch_clear(["A2:Z50"])  # Clear the Sheet from A2 to Z50
        in_progress_cameras = get_inprogress_cams(camera_location)
        update_progress(in_progress_cameras, camera_location,
                        worksheet, "quality")
        cameras.clear()


if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        print(str(e))
    finally:
        os.system('pause')
