###############################################################################
#       ANS DATA PARSER
#       Created by Alex Nellis
#       31-Mär-25
###############################################################################

import os
import pandas as pd

cwd = os.getcwd()
data_loc = cwd + '/data/'

def loadData():
    """
    Load all .xlsx and .csv files from the /data/ folder.
    -------
    Returns
        labels: list(str) 
            - list of strings of meeting names loaded in format of 'YEAR MEETING_NAME'

        contents: dict(str, np.ndarray(int64))
            - dictionary of meeting name keys and registered ANS member id data in int64

        member_list: dict(str, np.ndarray(int64))
            - dictionary of year keys (2022, 2023, 2024) and ANS member id data in int64
    """
    labels = []
    contents = {}
    member_list = {}
    
    for file in os.listdir(data_loc):
        f = os.path.join(data_loc,file)
        if os.path.isfile(f):
            title = f.split('/')[-1].split('-20')[-1][2:].split('.')[0]
            if 'member' not in f:
                labels.append(title)
                if f.endswith('.csv'):
                    contents[title] = pd.read_csv(f)['Record Number'].values
                elif f.endswith('.xlsx'):
                    xls = pd.ExcelFile(f)
                    temp = pd.read_excel(f, xls.sheet_names[-1])['Record Number'].values
                    contents[title] = temp
            else:
                if f.endswith('.csv'):
                    member_list[title] = pd.read_csv(f)['Record Number'].values
                elif f.endswith('.xlsx'):
                    xls = pd.ExcelFile(f)
                    member_list[title] = pd.read_excel(f, xls.sheet_names[-1])['Record Number'].values
    return labels, contents, member_list

def getYear(year, total_members_list, total_meetings_list):
    """
    Retrieve all data from the provided year
    -------
    Parameters
        year: int or str
            - year of data to retrieve eg. 2024 or '2024'

        total_members_list: dict(str, np.ndarray(int64))
            - dictionary containing all members data loaded from the /data/ folder

        total_meetings_list: dict(str, np.ndarray(int64))
            - dictionary containing all meeting and registered ANS member data loaded from the /data/ folder
    -------
    Returns
        member_list: np.ndarray(int64)
            - array containing all member data for the provided year

        meeting_list: dict(str, np.ndarray(int64))
            - dictionary containing all meeting and registered ANS member data for the provided year
    -------
    Raises
        KeyError
            - If no member list or meeting for the provided year is found
    """
    y = str(year) if type(year) is not str else year
    
    member_list = -1
    meeting_list = {}

    for key in total_members_list.keys():
        if y in key:
            member_list = total_members_list[key]
    for key in total_meetings_list.keys():
        if y in key:
            meeting_list[key] = total_meetings_list[key]
    
    if type(member_list) is int:
        raise KeyError(f'No member list for {year}')
    if len(meeting_list) == 0:
        raise KeyError(f'No meetings listed for {year}')
    return member_list, meeting_list

def getAttendance(member_ids, meetings):
    """
    Map attendance of each member to each meeting
    -------
    Parameters
        member_ids: np.ndarray(int64)
            - array of member id's to search for in the meeting data

        meetings: dict(str, np.ndarray(int64))
            - dictionary of meetings containing the attending ANS member data
    -------
    Returns
        meeting_long: pd.DataFrame
            - dataframe containing a boolean mapping of ANS members to the meetings they attended
    """
    meeting_log = pd.DataFrame(index = member_ids, columns = meetings.keys())
    meeting_log[meeting_log.isna()] = False
    
    for member in member_ids:
        for key in meetings.keys():
            if member in meetings[key]:
                meeting_log.at[member,key] = True
    return meeting_log

def totalAttendance(attendance, meetings, exclude = None):
    """
    Return the overlapping attendance of ANS members to the meetings provided
        This function will provide the number of members that attended at minimum the meetings provided
    -------
    Parameters
        attendance: pd.DataFrame
            - dataframe containing the boolean mapping of ANS members to the meetings they attended

        meetings: list(str)
            - list of meeting keys to search for member in format 'YEAR MEETING_NAME'

        exclude: list(str) or None
            - list of meeting keys to exclude attendance from in format 'YEAR MEETING_NAME'
                if a member attended one of these meetings, they will not be added to the result
    -------
    Returns
        result: pd.DataFrame
            - dataframe containing the ANS members who attended all of the searched meetings and none of the excluded meetings

        result.shape[0]: int
            - number of members in result
    -------
    Raises
        KeyError
            - If meetings and exclude both contain the same meeting (to prevent repetitive operation and potential input mistake)
    """
    mask = True
    for key in meetings:
        if exclude and key in exclude:
            raise KeyError(key + ' cannot be searched and removed simultaneously')
        mask = (mask) & (attendance[key])

    if exclude:
        for key in exclude:
            mask = (mask) & (not attendance[key])

    result = attendance[mask]
    return result, result.shape[0]

#"labels" contains the keys for meetings that can be used in totalAttendance
labels, contents, members = loadData()
attendance = getAttendance(*getYear(2024, members, contents))
result, total = totalAttendance(attendance, ['2024 NETS','2024 PBNC'])






    
