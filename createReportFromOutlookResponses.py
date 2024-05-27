# Built-in libraries
from os.path import dirname, join
from os import access, F_OK
from datetime import date
# External libraries
from win32com.client import CDispatch, Dispatch
import pandas as pd
from pandas import DataFrame, ExcelWriter
from response_record import ResponseRecord

def get_post_office_box(name_of_post_office_box: str) -> CDispatch:
    """Opens outlook application and a post office box which 
    is connected to the current account.
    
    Args:
        name_of_post_office_box (str)

    Returns:
        CDispatch: Represents a Folders object of outlook object modell.
    """    
    application: CDispatch = Dispatch('outlook.application')
    outlook_namespace: CDispatch = application.GetNamespace("MAPI")
    folders: CDispatch = outlook_namespace.Folders
    return folders[name_of_post_office_box]

def get_folder_of_inbox(parent_folder: CDispatch, folder_list: list[str]) -> CDispatch:
    """Starting from parent folder, the folders in post office box 
    will be searched for last folder name in passed folder list.
    The search order is equal to the order of string in passed folder list. 

    Args:
        parent_folder (CDispatch): Initial folder for searching.
        folder_list (list[str]): Contents folder names of post office box.

    Returns:
        CDispatch: Represents a Folder item of outlook object model.
    """    
    folder_name: str
    index: int
    folder: CDispatch
    for index, folder_name in enumerate(folder_list):
        if index == 0:
            folder = parent_folder.Folders[folder_name]
        else:
            folder = folder.Folders[folder_name]
    return folder
        
def item_is_response_to_invitation(item: CDispatch) -> bool:
    """Returns true if item is a positive, negative or tentative
    response to an invitation.

    Args:
        item (CDispatch): Either a MeetingItem or a MailItem of outlook object modell.

    Returns:
        bool: False if item.MessageClass is 'IPM.Note' (written mail) or 'IPM.Schedule.Meeting.Notification.Forward'
    """    
    message_class: str = item.MessageClass
    if message_class != 'IPM.Note' and message_class != 'IPM.Schedule.Meeting.Notification.Forward':
        return True

def get_dict_with_response_data(response: ResponseRecord) -> dict:
    """Passes information from passed object into dictionary.

    Args:
        response (ResponseRecord): Contents data from invitation for training 
        and information meetings due to DyNAMO project.

    Returns:
        dict: Contents the attributes
            - Name of training
            - Training date
            - First name of participiant
            - Last name of participiant
            - ID of participiant
            - Kind of response (positive, negative or tentative)
    """    
    id: str = response.get_id_of_participant
    first_name: str = response.participant_first_name
    last_name: str = response.participant_last_name
    training: str = response.name_of_training
    training_date: date = response.training_date
    response_type: str = response.response_to_invitation
    header: tuple[str] = get_header()
    temp_dict: dict = {header[0]: training,
                       header[1]: training_date,
                       header[2]: first_name,
                       header[3]: last_name,
                       header[4]: id,
                       header[5]: response_type}
    return temp_dict

def get_header() -> tuple[str]:
    return ('Schulung',
            'Datum',
            'Vorname',
            'Nachname',
            'ID',
            'Antwort')

def get_dataframe_from_folder_items(outlook_folder: CDispatch) -> DataFrame:
    """Passes the following information from invitation responses in a dataframe:
    
    - personal ID, first and last name of participiant
    - Conversation date and topic
    - Timestamp of response

    Args:
        outlook_folder (CDispatch): A folder of a post office box.

    Returns:
        DataFrame
    """    
    response: ResponseRecord
    response_list: list[dict] = []
    item: CDispatch
    
    for item in outlook_folder:
        if item_is_response_to_invitation(item):
            response = ResponseRecord()
            response.import_meeting_item(item)
            response_list.append(get_dict_with_response_data(response))        
        # Close MeetingItem. Changes to the document are discarded.
        item.close(1)

    return DataFrame(data=response_list)    

def export_data_to_report_file(extracted_data_from_outlook: DataFrame) -> None:
    """The passed dataframe will be adapted to table in report file.
    If file does not exists, it will be created in current folder.

    Args:
        extracted_data_from_outlook (DataFrame): 
        The dataframe has to be the same columns like the table in report file.
    """       
    
    def add_data_to_table_in_report_file() -> None:
        """Appends new data to existing data in report file and removes duplicates.
        """        
        with ExcelWriter(path=FILE_PATH, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer: # pylint: disable=abstract-class-instantiated
            existing_data: DataFrame = pd.read_excel(FILE_PATH, sheet_name=SHEET_NAME)
            summary: DataFrame = pd.concat([existing_data, extracted_data_from_outlook])
            summary.drop_duplicates(inplace=True)
            summary.to_excel(writer,sheet_name=SHEET_NAME, index=False)
        
    CURRENT_DIR = dirname(__file__)
    FILE_NAME: str = 'response_report.xlsx'
    FILE_PATH: str = join(CURRENT_DIR, FILE_NAME)
    SHEET_NAME: str = 'response_report'
    
    # If path exists, sheet in workbook will be replaced.
    if access(path=FILE_PATH,mode=F_OK):
        add_data_to_table_in_report_file()
    else:
        with ExcelWriter(path=FILE_PATH, engine='openpyxl') as writer: # pylint: disable=abstract-class-instantiated 
            data.to_excel(writer,sheet_name=SHEET_NAME, index=False)
            
if __name__ == '__main__':    
    # A full list of OIDefaultFolders enumeration can be found on https://docs.microsoft.com/en-us/office/vba/api/outlook.oldefaultfolders
    parent_folder: CDispatch = get_post_office_box(name_of_post_office_box='my.mailaccount@example.com')
    test_folder: CDispatch = get_folder_of_inbox(parent_folder=parent_folder, folder_list=['Posteingang','Test'])
    content_of_folder: CDispatch = test_folder.Items
    data: DataFrame = get_dataframe_from_folder_items(content_of_folder)

    del parent_folder, test_folder, content_of_folder
    export_data_to_report_file(data)
