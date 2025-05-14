###############################################################################
#                            ANS DATA PARSER                                  #
#                         Created by Alex Nellis                              #
#                               31-Mär-25                                     #
#                         Edited by Luiz Aldeia                               #
#                               01-Abr-25                                     #
###############################################################################

import os
import pandas as pd
import numpy as np
import holoviews as hv
from holoviews import dim

# Enable the Bokeh extension for Holoviews (this allows interactive plots in the browser)
hv.extension('bokeh')

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
        #some files uses 'Record Number' and others 'Account: ANS ID', this will take care
        #of this problem when accessing the member ID
        f = os.path.join(data_loc,file)
        if os.path.isfile(f):
            title = f.split('/')[-1].split('-20')[-1][2:].split('.')[0]
            if 'member' not in f:
                labels.append(title)
                if f.endswith('.csv'):
                    df = pd.read_csv(f)
                elif f.endswith('.xlsx'):
                    xls = pd.ExcelFile(f)
                    df = pd.read_excel(f, xls.sheet_names[-1])
                else:
                    continue
                
                if 'Record Number' in df.columns:
                    contents[title] = df['Record Number'].values
                elif 'Account: ANS ID' in df.columns:
                    contents[title] = df['Account: ANS ID'].values
            else:
                if f.endswith('.csv'):
                    df = pd.read_csv(f)
                elif f.endswith('.xlsx'):
                    xls = pd.ExcelFile(f)
                    df = pd.read_excel(f, xls.sheet_names[-1])
                else:
                    continue
                
                if 'Record Number' in df.columns:
                    member_list[title] = df['Record Number'].values
                elif 'Account: ANS ID' in df.columns:
                    member_list[title] = df['Account: ANS ID'].values
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

def generateChordDiagram(attendance, selected_meetings, year):
    """
    Generate an interactive chord diagram showing the correlation between the selected meetings
    and save it as an HTML file.
    """

    meetings_names = [meeting.split(' ', 1)[1] for meeting in selected_meetings if ' ' in meeting]
    

    overlap_matrix = np.zeros((len(meetings_names), len(meetings_names)))
    overlap_diag = np.zeros(len(meetings_names))

    for i, meeting_i in enumerate(meetings_names):
        for j, meeting_j in enumerate(meetings_names):
            column_i = f"{year} {meeting_i}"
            column_j = f"{year} {meeting_j}"
            if i == j:
                overlap_diag[i] = (attendance[column_i]).sum()
                continue

            # Check if both columns exist before proceeding
            if column_i in attendance.columns and column_j in attendance.columns:
                overlap = (attendance[column_i] & attendance[column_j]).sum()
                overlap_matrix[i, j] = overlap
            else:
                print(f"Warning: Columns {column_i} or {column_j} not found in the attendance data for {year}. Skipping pair.")


    links = []
    for i in range(len(meetings_names)):
        for j in range(i + 1, len(meetings_names)):
            if overlap_matrix[i, j] > 0:
                links.append({
                    'source': meetings_names[i], 
                    'target': meetings_names[j], 
                    'value': overlap_matrix[i, j]
                })
    
    links_df = pd.DataFrame(links)

    chord = hv.Chord(links_df).opts(
        labels='index', 
        cmap='Category10',
        edge_cmap='Category10',
        edge_color=dim('source').str(),  
        node_color=dim('index').str(), 
        width=800, 
        height=800,  
        node_radius=0.025,  
        padding=1,  
        title=f"{year} Meeting Correlations",  
        label_text_font_size='12pt',  
        edge_line_width=3 
    )
    
    hv.save(chord, f'chord_diagram_overlap_{year}.html')
    print(f"Chord diagram saved as 'chord_diagram_{year}.html'")
    return overlap_matrix + np.diag(overlap_diag)




if __name__ == "__main__":
    while True:
        year = input("\nEnter the year to be analyzed: ").strip()
        labels, contents, members = loadData()
        available_meetings = [label.replace(year, '').strip() for label in labels if year in label]
        
        if available_meetings:
            break
        else:
            print(f"No meeting data found for {year}. Please enter a valid year.")
    
    while True:
        print("\nThe following meeting data are available:\n\t" + " ".join(available_meetings) + " ALL\n")
        meetings_input = input("Which meeting(s) should be analyzed (separate by commas): ").strip().upper()
        
        if meetings_input == "ALL":
            selected_meetings = [f"{year} {m}" for m in available_meetings]
            break
        else:
            selected_meetings = [f"{year} {m.strip().upper()}" for m in meetings_input.split(',') if m.strip().upper() in map(str.upper, available_meetings)]
        
        invalid_meetings = [m.strip().upper() for m in meetings_input.split(',') if m.strip().upper() not in map(str.upper, available_meetings)]
        
        if not invalid_meetings:
            break
        else:
            print(f"\nThe following meetings are not on the list: {', '.join(invalid_meetings)}\nPlease enter the meeting names again.")
    
    print("\nSelected conferences:", selected_meetings)
    
    attendance = getAttendance(*getYear(year, members, contents))
    
    attendance_filename = f"{year}_attendance.csv"
    attendance.to_csv(attendance_filename, index=True)
    print(f"Attendance data saved to {attendance_filename}\n")
    
    result, total = totalAttendance(attendance, selected_meetings)

    print(total," members participated in all the selected meetings\n")
    print(result)
    
    if total > 0:
        output_filename = input("\nEnter the output file name for the results (including .csv extension): ").strip()
        result.to_csv(output_filename, index=True)
        print(f"Results saved to {output_filename}")

    if meetings_input == "ALL":
        print("\nGenerating chord diagram for meeting correlations...")
        matrix = generateChordDiagram(attendance, selected_meetings, year)
        for i in range(len(matrix)):
            matrix[i] /= matrix[i,i]
        matrix *= 100
        table_filename = f"{year}_overlap_table.csv"
        overlap_table = pd.DataFrame(data = matrix, index = selected_meetings, columns = selected_meetings)
        overlap_table.to_csv(table_filename, index = True)
        print(f"Overlap Table data saved to {table_filename}")
        
        overlap = attendance[attendance.sum(axis=1) > 1]
        overlap_filename = f"{year}_attendance_overlap.csv"
        overlap.to_csv(overlap_filename, index=True)
        print(f"Overlap attendance data saved to {overlap_filename}")
