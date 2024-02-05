from customtkinter import *
from tkinter import filedialog
import pandas as pd
from PIL import Image
from typing import List
import requests
import threading

default_theme = 0

set_appearance_mode("dark")
set_default_color_theme("dark-blue")

icon_image_path = ".\\assets\\main_icon.ico"
light_mode_image_path = ".\\assets\\light_mode.png"
dark_mode_image_path = ".\\assets\\dark_mode.png"

version = 1.6

root = CTk()
root.resizable(False, False)
root.title(f"EID Tool v{version}")
root.iconbitmap(icon_image_path)

theme_image = CTkImage(light_image=Image.open(dark_mode_image_path), dark_image=Image.open(light_mode_image_path), size=(10, 10))

window_height = 500
popup_height = 100
window_width, popup_width = 400, 500


screen_width = root.winfo_screenwidth()
screen_height = root.winfo_screenheight()

x_cordinate = int((screen_width/2) - (window_width/2))
y_cordinate = int((screen_height/2) - (window_height/2))

root.geometry("{}x{}+{}+{}".format(window_width, window_height, x_cordinate, y_cordinate))

market_mapping = (
    ("500 Meg", "[C, E]"),
    ("All Digital 500Meg", "[C, E]"),
    ("All Digital 400Meg", "[A, B]"),
    ("LIMITED", "[J, Q]"),
    ("GIG - FIBER", "[G,  F]"),
    ("PREGIG", "[K, O]"),
    ("GIG FIBER DATA ONLY", "[G, F]"),
    ("VIDEO ONLY", "[V]"),
    ("GIG", "[M, N]"),
    ("Gig", "[M, N]"),
    )


def check_for_updates():
    global update_label
    url = "https://raw.githubusercontent.com/ganesh2257767/EID-Tool/main/update_changelogs.txt"
    response = requests.get(url)
    if response.status_code == 200:
        latest_version = float(response.text.split("\n")[0].split("v")[-1])
        if version < latest_version:
            update_label.configure(text=f"New version v{latest_version} available.", text_color="red")
            return
        if version >= latest_version:
            update_label.configure(text=f"All good, currently on the latest version v{latest_version}.", text_color="green")
            return
    else:
        update_label.configure(text=f"Unable to check latest version details at this time.", text_color="red")
        return


def handle_error_popups(err_message: str) -> None:
    """
        Handles all exceptions with a popup and an appropriate error message.
        
        Parameters:
            err_message: str -> Takes in a string error message to display on the popup window.
    """
    popup = CTkToplevel(root)
    popup.title("Error")
    popup.iconbitmap(icon_image_path)
    popup.geometry("{}x{}+{}+{}".format(popup_width, popup_height, x_cordinate-50, y_cordinate+100))
    popup.geometry("500x150")
    popup.grab_set()
    label = CTkLabel(popup, text=err_message, padx=10, pady=10)
    label.pack(expand=True)
    button = CTkButton(popup, text="Dismiss", fg_color="red", hover_color="#d9453b", command=lambda: popup.destroy())
    button.pack(padx=10, pady=10)
    popup.bind('<Return>', lambda x: popup.destroy())
    popup.bind('<Escape>', lambda x: popup.destroy())
    popup.bind('<Key-space>', lambda x: popup.destroy())
    
    popup.wm_transient(root)


def format_for_table(lst: List) -> List:
    """
    format_for_table - Takes in a list and returns a list of list for displaying

    Takes in a list of values and converts it into a list of lists with 6 elements each (except the last one in some ocassions).
    Helps in displaying it on the result frame cleanly.

    :param lst:  - A list of values from the dataframe.
    :type lst: List
    :return: temp - Formatted list of lists with 6 elements in each sub-list.
    :rtype: List
    """
    temp = []

    for i in range(0, len(lst), 6):
        temp.append(lst[i:i+6])

    return temp


master_matrix_dataframe = None
def get_master_matrix() -> None:
    """
    get_master_matrix - Uploads the Master Matrix file supplied and creates a Pandas DataFrame.

    Uploads the Master Matrix file supplied and creates a Pandas DataFrame, also throws exceptions if the file format is not excel (.xlsx)
    """
    global upload_master_indicator, frame2, master_matrix_dataframe

    master_matrix_path = filedialog.askopenfilename()
    read_cols = ['Corp', 'CONCATENATE', 'RESI EID', 'Market', 'Altice One']
    try:
        master_matrix_dataframe = pd.read_excel(master_matrix_path, usecols=read_cols, converters={'Corp':str, 'CONCATENATE': str})
        for market in market_mapping:
            master_matrix_dataframe = master_matrix_dataframe.replace(*market, regex=True)
    except ImportError as e:
        frame2.pack_forget()
        handle_error_popups(e)
    except FileNotFoundError:
        pass
    except ValueError as e:
        frame2.pack_forget()
        upload_master_indicator.configure(text=f"No file uploaded", text_color="red")
        if "Excel file format cannot be determined, you must specify an engine manually" in e.args[0]:
            handle_error_popups("Please upload an excel file only!")
        elif "Usecols do not match columns, columns expected but not found" in e.args[0]:
            error_column_name = str(e)[str(e).index("'"):].replace(']', '')
            handle_error_popups(f"Column(s) {error_column_name} not found.")

    else:
        upload_master_indicator.configure(text=f"File {master_matrix_path.split('/')[-1]} uploaded", text_color="green")
        frame2.pack(padx=10, pady=10, fill="both")


def get_corp_ftax_from_offer_id(env: str, offer_id: str) -> None:
    """
    get_corp_ftax_from_offer_id - Displays the corp and ftax combination depending on the environment and offer ID passed. 

    Takes environment and offer ID and displays the corp and ftax combinations as well as Market type supported in that corp-ftax combination.

    :param env: Environment for which the combinations are requested. Values passed from th GUI to avoid incorrect values.
    :type env: str
    :param offer_id: Offer ID to check the corp-ftax combination.
    :type offer_id: str
    """
    if not all((env, offer_id)):
        handle_error_popups("Won't work if there's missing values!")
        return

    match env:
        case "QA INT":
            corp = ['7702', '7704', '7710', '7715']
        case "QA 1":
            corp = ['7708', '7711']
        case "QA 2":
            corp = ['7712', '7709']
        case "QA 3":
            corp = ['7707', '7714']
        case _:
            corp = ['7701', '7703', '7705', '7706', '7713']

    corpftax_altice_list = []
    corpftax_legacy_list = []
    smb_list = []

    offer_eid = {eid_dataframe['ELIGIBILITY_ID'][i] for i in eid_dataframe.index if eid_dataframe['OFFER_ID'][i] == offer_id}
    for j in master_matrix_dataframe.index:
        if master_matrix_dataframe['Corp'][j] in corp:
            if master_matrix_dataframe['Altice One'][j] == 'Y' and master_matrix_dataframe['RESI EID'][j] in offer_eid:
                f = f"{master_matrix_dataframe['CONCATENATE'][j][4:]:0>2}"
                corpftax_altice_list.append(
                    f"{master_matrix_dataframe['CONCATENATE'][j][:4]}-{f} - {master_matrix_dataframe['Market'][j].strip()} - {master_matrix_dataframe['RESI EID'][j].strip()}")

            elif master_matrix_dataframe['Altice One'][j] == 'N' and master_matrix_dataframe['RESI EID'][j] in offer_eid:
                f = f"{master_matrix_dataframe['CONCATENATE'][j][4:]:0>2}"
                corpftax_legacy_list.append(
                    f"{master_matrix_dataframe['CONCATENATE'][j][:4]}-{f} - {master_matrix_dataframe['Market'][j].strip()} - {master_matrix_dataframe['RESI EID'][j].strip()}")

            elif master_matrix_dataframe['RESI EID'][j] in offer_eid:
                f = f"{master_matrix_dataframe['CONCATENATE'][j][4:]:0>2}"
                smb_list.append(f"{master_matrix_dataframe['CONCATENATE'][j][:4]}-{f} - {master_matrix_dataframe['Market'][j].strip()} - {master_matrix_dataframe['RESI EID'][j].strip()}")


    if corpftax_legacy_list or corpftax_altice_list:
        result_altice = format_for_table(corpftax_altice_list)
        result_legacy = format_for_table(corpftax_legacy_list)
        display_result_table(result_altice, f"Altice Combinations for Offer {offer_id}", offer_id, result_legacy, f"Legacy Combinations for Offer {offer_id}")

    elif smb_list:
        result_smb = format_for_table(smb_list)
        display_result_table(result_smb, 'SMB Combinations')

    else:
        handle_error_popups(f'Offer {offer_id} not available in {corp} or is invalid!\n\nPlease check Offer ID or change corp and try again!')



eid_dataframe = None
def get_eid_sheet() -> None:
    """
    get_eid_sheet - Uploads the EID file supplied and creates a Pandas DataFrame.

    Uploads the EID file supplied and creates a Pandas DataFrame, also throws exceptions if the file format is not excel (.xlsx)
    """

    global upload_eid_indicator, frame2, eid_dataframe

    eid_path = filedialog.askopenfilename()
    try:
        eid_dataframe = pd.read_excel(eid_path, usecols=['ELIGIBILITY_ID', 'OFFER_ID'], converters={'ELIGIBILITY_ID': str, 'OFFER_ID': str})
    except ImportError as e:
        frame5.pack_forget()
        handle_error_popups(e)
    except FileNotFoundError:
        pass
    except ValueError as e:
        frame5.pack_forget()
        upload_eid_indicator.configure(text=f"No file uploaded", text_color="red")
        if "Excel file format cannot be determined, you must specify an engine manually" in e.args[0]:
            handle_error_popups("Please upload an excel file only!")
        elif "Usecols do not match columns, columns expected but not found" in e.args[0]:
            error_column_name = str(e)[str(e).index("'"):].replace(']', '')
            handle_error_popups(f"Column(s) {error_column_name} not found.")
    else:
        upload_eid_indicator.configure(text=f"File {eid_path.split('/')[-1]} uploaded", text_color="green")
        frame5.pack(padx=10, pady=10, fill="both")
        oid_input.focus_set()


def from_eid(eid: str) -> None:
    """
    from_eid - Displays Corp and Ftax combination for provided EID value.

    Displays the Corp and Ftax combination for a provided EID.

    :param eid: EID to check the Corp-Ftax combination.
    :type eid: str
    """    
    if not eid:
        handle_error_popups('You need to pass in an EID value for this to work!')
        return

    corp_ftax = []
    for i in master_matrix_dataframe.index:
        if master_matrix_dataframe['RESI EID'][i] == eid:
            f = f"{master_matrix_dataframe['CONCATENATE'][i][4:]:0>2}"
            a = f"""{master_matrix_dataframe['CONCATENATE'][i][:4]}-{f} - {master_matrix_dataframe['Market'][i].strip()}"""
            corp_ftax.append(a)

    if corp_ftax:
        result = format_for_table(corp_ftax)

        display_result_table(result, f"Corp-Ftax - Market combinations for {eid}", eid)
    else:
        handle_error_popups(f'{eid} not available or invalid, please check and try again!')


def display_result_table(result1: List, heading1: str, title:str, result2: List=None, heading2: str=None) -> None:
    """
    display_result_table - Creates a niceley styled result table.

    Takes in the result and heading and creates a nicely styled and padded display table from a combination of Text and Frame widgets.

    :param result1: [Required] - List of list of the result values to be displayed.
    :type result1: List
    :param heading1: [Required] - Heading for the result1 data set.
    :type heading1: str
    :param result2: [Optional] - To be passed in case the Offer ID has Altice and Legacy combinations, defaults to None
    :type result2: List, optional
    :param heading2: [Optional] - To be passed in case the Offer ID has Altice and Legacy combinations, defaults to None
    :type heading2: str, optional
    """    
    result_popup = CTkToplevel(root)
    result_popup.title(f"Result for {title}")
    result_popup.iconbitmap(icon_image_path)
    result_popup.geometry("{}+{}".format(x_cordinate-600, y_cordinate-150))
    result_popup.wm_transient(root)
    result_popup.grab_set()

    min_width = 350
    
    if result1:
        rows = len(result1)
        try:
            columns = len(result1[0])
        except IndexError:
            columns = 1
        result_frame1 = CTkScrollableFrame(result_popup, height=100, corner_radius=0, fg_color="transparent")
        result_frame1.grid(row=0, column=0, padx=20, pady=20, sticky="nsew")
        
        label = CTkLabel(result_frame1, text=heading1, font=('Segoe UI', 18, 'bold'))
        label.grid(row=0, columnspan=6, padx=10, pady=10)
        
        button_row_start = len(result1)
        
        total_width = 0

        for i in range(rows):
            for j in range(columns):
                try:
                    e = CTkEntry(result_frame1, width=len(result1[i][j])*9, corner_radius=0, justify=CENTER)
                except IndexError:
                    continue
                try:
                    e.grid(row=i+1, column=j, padx=10, pady=10)
                    e.insert(END, result1[i][j])
                    e.configure(state="disabled")
                except IndexError:
                    e.grid_forget()
                    break
                if i == 0:
                    total_width += len(result1[i][j])*9 + 30

        result_frame1.configure(width=max(total_width, min_width))

    if result2:
        rows = len(result2)
        try:
            columns = len(result2[0])
        except IndexError:
            columns = 1
        result_frame2 = CTkScrollableFrame(result_popup, height=100, corner_radius=0, fg_color="transparent")
        result_frame2.grid(row=1, column=0, padx=20, pady=20, sticky="nsew")

        label = CTkLabel(result_frame2, text=heading2, font=('Segoe UI', 18, 'bold'))
        label.grid(row=0, columnspan=6, padx=10, pady=10)
        
        button_row_start = len(result2) + len(result1) + 1
        
        total_width = 0
        
        for i in range(rows):
            for j in range(columns):
                try:
                    e = CTkEntry(result_frame2, width=len(result2[i][j])*9, corner_radius=0, justify=CENTER)
                except IndexError:
                    continue
                try:
                    e.grid(row=i+1, column=j, padx=10, pady=9)
                    e.insert(END, result2[i][j])
                    e.configure(state="disabled")
                except IndexError:
                    e.grid_forget()
                    break
                if i == 0:
                    total_width += len(result2[i][j])*9 + 30

        result_frame2.configure(width=max(total_width, min_width))

    button = CTkButton(result_popup, text="Close", fg_color="red", hover_color="#d9453b", command=lambda: result_popup.destroy())
    button.grid(row=button_row_start, columnspan=6, padx=10, pady=10)
    result_popup.bind('<Return>', lambda x: result_popup.destroy())
    result_popup.bind('<Escape>', lambda x: result_popup.destroy())
    result_popup.bind('<Key-space>', lambda x: result_popup.destroy())
    # result_popup.wm_transient(root)


def get_radio_value() -> None:
    """
     Switches the window based on the radio button selected.
    """

    val = radio_selection.get()
    if val == 1:
        frame3.grid(row=1, column=0, padx=10, pady=10, columnspan=2, sticky=N+S+W+E)
        eid_input.focus()
        frame4.grid_forget()
    else:
        frame4.grid(row=1, column=0, padx=10, pady=10, columnspan=2, sticky=E+W)
        frame3.grid_forget()


def change_theme() -> None:
    """
        Switches between light and dark themes.
    """

    global default_theme
    if default_theme == 0:
        default_theme = 1
        set_appearance_mode("light")
    else:
        default_theme = 0
        set_appearance_mode("dark")


# Frame 1 (Button to upload and display uploaded Master Matrix)
frame1 = CTkFrame(root)
frame1.pack(padx=10, pady=10, fill="both")

# Frame 1 widgets and variable
upload_matrix_button = CTkButton(frame1, text="Select the latest Master Matrix", command=get_master_matrix)

upload_matrix_button.pack(pady=(20, 5))

upload_master_indicator = CTkLabel(frame1, text="No file uploaded", text_color="red")
upload_master_indicator.pack(pady=(5, 10))

# Frame 2 (Radio buttons to select the search criteria, either with EID or Offer ID) 1630941  
frame2 = CTkFrame(root)

# Frame 2 widgets
radio_selection = IntVar()

search_with_eid = CTkRadioButton(frame2, text="Search with EID", variable=radio_selection, value=1, command=get_radio_value)
search_with_eid.grid(row=0, column=0, padx=(10, 0), pady=10)

search_with_oid = CTkRadioButton(frame2, text="Search with Offer ID", variable=radio_selection, value=2, command=get_radio_value)
search_with_oid.grid(row=0, column=1, padx=(0, 10), pady=10)
radio_selection.set(1)

# Frame 3 (Input box and submit button for EID search criteria)
frame3 = CTkFrame(frame2)
frame3.grid(row=1, column=0, padx=10, pady=10, columnspan=2, sticky=N+S+E+W)

frame2.grid_columnconfigure(0, weight=1)
frame2.grid_rowconfigure(0, weight=1)

# Frame 3 widgets
eid_var = StringVar()

eid_label = CTkLabel(frame3, text="Enter EID")
eid_label.pack(padx=10, pady=(10, 1))

eid_input = CTkEntry(frame3, width=200, textvariable=eid_var, takefocus=True)
eid_input.pack(padx=10, pady=(1, 5))
eid_input.focus_set()

eid_var.trace('w', lambda *args: eid_var.set(eid_var.get().upper().strip()))

eid_submit = CTkButton(frame3, text="Submit", width=75, command=lambda:from_eid(eid_var.get()))
eid_submit.pack(padx=10, pady=(5, 10))

eid_input.bind('<Return>', lambda x:from_eid(eid_var.get()))
eid_input.bind('<Key-space>', lambda x:from_eid(eid_var.get()))

# Frame 4 (Button to upload EID sheet and display uploaded sheet)
frame4 = CTkFrame(frame2)

frame2.grid_columnconfigure(0, weight=1)
frame2.grid_rowconfigure(0, weight=1)

# Frame 4 widgets
eid_dataframe = None
upload_eid_sheet = CTkButton(frame4, text="Select the latest EID Sheet", command=get_eid_sheet)
upload_eid_sheet.pack(padx=10, pady=(20, 1))

upload_eid_indicator = CTkLabel(frame4, text="No file uploaded", text_color="red")
upload_eid_indicator.pack(padx=10, pady=(1, 5))

# Frame 5 (Input, Dropdown and Submit button to search with Offer ID)
frame5 = CTkFrame(frame4)

# Frame 5 widgets
oid_label = CTkLabel(frame5, text="Enter Offer ID")
oid_label.grid(row=0, column=0, padx=(20, 10), pady=(5, 5))

oid_var = StringVar()

oid_input = CTkEntry(frame5, textvariable=oid_var)
oid_input.grid(row=1, column=0, padx=(20, 10), pady=(5, 5))

oid_var.trace('w', lambda *args: oid_var.set(oid_var.get().strip()))

env_label = CTkLabel(frame5, text="Select Env.")
env_label.grid(row=0, column=1, padx=10, pady=(5, 5))

env_var = StringVar(value="QA INT")
env_drop = CTkOptionMenu(frame5, values=["QA INT", "QA 1", "QA 2", "QA 3", "Others"], variable=env_var)
env_drop.grid(row=1, column=1, padx=10, pady=(5, 5))

oid_submit = CTkButton(frame5, text="Submit", width=75, command=lambda: get_corp_ftax_from_offer_id(env_var.get(), oid_var.get()))
oid_submit.grid(row=2, column=0, columnspan=2, padx=(20, 10), pady=(5, 10), sticky=E+W)

oid_input.bind('<Return>', lambda x: get_corp_ftax_from_offer_id(env_var.get(), oid_var.get()))
oid_input.bind('<Key-space>', lambda x: get_corp_ftax_from_offer_id(env_var.get(), oid_var.get()))

# Frame 6 (Footer section with Tool version and theme change mechanism)
frame6 = CTkFrame(root, height=50)
frame6.pack(side=BOTTOM, padx=10, pady=(0, 10), fill=X)

# Frame 6 widgets
version_label = CTkLabel(frame6, text=f"EID Tool v{version}", width=20)
version_label.pack(side=LEFT, padx=5, pady=5)

light_dark_button = CTkButton(frame6, image=theme_image, text="", width=16, height=16, fg_color="white", command=change_theme)
light_dark_button.pack(side=RIGHT, padx=(0, 20), pady=5)

theme_label = CTkLabel(frame6, text="Change theme", width=25)
theme_label.pack(side=RIGHT, padx=5, pady=5)

# Update section on root
update_label = CTkLabel(root, text="Checking for updates in the background...")
update_label.pack(side=BOTTOM)

if __name__ == "__main__":
    threading.Thread(target=check_for_updates).start()
    root.mainloop(0)