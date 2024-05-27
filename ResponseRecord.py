from datetime import date
from typing import List
from win32com.client import CDispatch

class ResponseRecord:
    """Class for extracting and returning information from MeetingItem objects 
    of outlook object modell.
    
    Input: A CDispatch object which represents a response 
    to an invitation. 
    """    
    def __init__(self) -> None:
        self._training_date: date = None
        self._name_of_training: str = None
        self._record_timestamp: date = None
        self._message_class: str = None
        self._participant_name: str = None
        self._ldap_data_of_sender_mailaddress: str = None
        
    def import_meeting_item(self, meeting_item: CDispatch) -> None:
        """Extracts information from MeetingItem.

        Further information of MeetingItem can be found on 
        https://docs.microsoft.com/en-us/office/vba/api/Outlook.MeetingItem
        
        Args:
            meeting_item (CDispatch): Represents a MeetingItem.
        """        
        self._training_date = meeting_item.ReminderTime
        self._name_of_training = meeting_item.ConversationTopic
        self._message_class = meeting_item.MessageClass
        # SenderName property returns the separated name like 'Mustermann, Max'
        self._participant_name = meeting_item.SenderName
        self._ldap_data_of_sender_mailaddress = meeting_item.SenderEmailAddress
    
    @property
    def get_id_of_participant(self) -> str:
        """This property extracts CN from the SenderEmailAddress property, which looks like
        
        '/o=Company/ou=Exchange Administrative Group (XYZ134)/cn=Recipients/cn=[*xy-id]'
        
        See further information of SenderEmailAddress on 
        https://docs.microsoft.com/en-us/previous-versions/office/developer/office-2003/aa171942%28v=office.11%29
        
        Returns:
            [str]: Unique ID of an employee
        """        
        
        def extract_id_from_common_name(common_name: str) -> str:
            """Assumption that the ID begins with xy000.
            """
            LENGTH_OF_ID: int = 9
            position_of_id: int = common_name.find('xy000')
            return common_name[position_of_id:position_of_id + LENGTH_OF_ID]
        
        def get_common_name_with_id() -> str:
            split: List[str] = []
            ldap_data: str = self._ldap_data_of_sender_mailaddress
            # The string will be formated to lower case for preventing
            # errors because of capitalized identifier like 'CN'.
            ldap_data = ldap_data.lower()
            split = ldap_data.split('cn=')
            return split[2]  
        
        common_name_with_id: str = get_common_name_with_id()
        return extract_id_from_common_name(common_name_with_id)
    
    @property
    def training_date(self) -> str:
        date_format: str = '%d.%m.%Y'
        return self._training_date.strftime(date_format)
    
    @property
    def participant_first_name(self) -> str:
        substring: str
        substring = self._participant_name.split(sep=',')[1]
        return substring.strip()
    
    @property
    def participant_last_name(self) -> str:
        return self._participant_name.split(sep=',')[0]
    
    @property
    def name_of_training(self) -> str:
        return self._name_of_training.replace('WG: ', '')
    
    @property
    def response_to_invitation(self) -> str:
        return self._get_response(self._message_class)
        
    def _get_response(self, message_class: str) -> str:
        """This function translates the value of MessageClass property.
        
        In case of MeetingItem objects, the following values are possible:
        
        - IPM.Schedule.Meeting.Resp.Pos = Positive response
        - IPM.Schedule.Meeting.Resp.Neg = Negative response
        - IPM.Schedule.Meeting.Resp.Tent = Tentative response
        - IPM.Schedule.Meeting.Notification.Forward = Invitation was forwarded

        Args:
            message_class (str): [Property of Outlook classes 'MeetingItem' or 'Action']

        Returns:
            str: [German translation of MessageClass]
        """        
        if message_class == 'IPM.Schedule.Meeting.Resp.Pos':
            return 'Zusage'
        elif message_class == 'IPM.Schedule.Meeting.Resp.Neg':
            return 'Absage'
        elif message_class == 'IPM.Schedule.Meeting.Resp.Tent':
            return 'Vorbehalt'
