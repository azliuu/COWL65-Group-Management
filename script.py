from itertools import cycle
import pandas as pd
import streamlit as stimport sys
import subprocess

# Install openpyxl if not installed
try:
    import openpyxl
except ImportError:
    subprocess.check_call([sys.executable, "-m", "pip", "install", "openpyxl"])
    import openpyxl  # Now try to import again


# from st_agrid import AgGrid

st.title("COWL 65 Group Generator")
file = st.file_uploader("Upload file", type=["xlsx"])


# Read the Excel sheet into a pandas DataFrame
if file is not None:
    df = pd.read_excel(file)

    # # Add an empty "Past Leaders" column to the DataFrame
    # df['Past Leaders'] = ''

    # Sort the DataFrame by Section and first name
    # print("this is timeslot", df['Section'])
    # print(df.sort_values(['Section'], inplace=True))
    df.sort_values(['Section'], inplace=True)

    # if leader column doesn't exist create leader column
    if 'Leader' not in df.columns:
        df['Leader'] = 0

    for index, row in df.iterrows():
        if (row['Leader'] == 1):
            df.at[index, 'Leader'] = -1

    # Group the DataFrame by Section
    grouped = df.groupby('Section')

    # Create a new DataFrame to store the groups
    groups = []

    # Iterate over each Section
    for time_slot, time_slot_data in grouped:
        # Shuffle the Section data to randomize the order
        time_slot_data = time_slot_data.sample(frac=1).reset_index(drop=True)

        # Calculate the number of groups needed
        num_groups = len(time_slot_data) // 4

        # Calculate how many people are leftover
        remainder = len(time_slot_data) % 4

        # Create a cycle of unique leaders for each group
        leaders = cycle(time_slot_data['Email Address'])

        # Create a set to track used leaders for this Section
        used_leaders = set()

        # Iterate over each group
        for i in range(num_groups):
            # Get the people for the current group
            group_people = time_slot_data.iloc[i * 4:(i + 1) * 4]

            # Check if a leftover person needs to be added to a group
            leftover = 0
            if i < remainder:
                leftover_person = time_slot_data.iloc[len(
                    time_slot_data) - i - 1: len(time_slot_data) - i]
                group_people = pd.concat([group_people, leftover_person])
                leftover = 1

        # Find a unique leader for the group
            leader = group_people.iloc[0]['Email Address']
            # print("leader: ")
            # print(leader)
            # print("group people: ")
            # print(group_people)
            # print(group_people[group_people['Leader']])
            # print("gp e")
            # print(group_people['Email Address Address'][0])
            # while (
            # leader in group_people['Email Address'].values or
            # leader in used_leaders or
            # (df.loc[df['Email Address'] == leader, 'Leader'].any() == -1)
            # ):
            # If the currently selected leader is marked as '-1',
            # try to find a member marked as '0' to be the new leader.
            # print("hihi")
            # print(df.loc[df['Email Address'] == leader, 'Leader'].any())
            # print("row")
            # print(df.loc[df['Email Address'] == leader])
            # print("leader")
            # # print(df.loc[df['Leader']])
            if df.loc[df['Email Address'] == leader, 'Leader'].any():
                new_leader_candidates = group_people[group_people['Leader'] == 0]
                if not new_leader_candidates.empty:
                    leader = new_leader_candidates.iloc[0]['Email Address']
                else:
                    # If no '0' leader candidates are available, select a member marked as '-1'.
                    leader = group_people[group_people['Leader']
                                          == -1].iloc[0]['Email Address']
            else:
                # If the currently selected leader is not marked as '-1',
                # select a new leader from available members.
                leader = group_people[group_people['Leader']
                                      == 0].iloc[0]['Email Address']
                # print(leader)
            # Set the leader status for the assigned leader
            df.loc[df['Email Address'] == leader, 'Leader'] = 1

            group_data = {
                'Section': time_slot,
                'Group': i + 1,
                # choose the first member of the group as the leader if the leader is not -1.
                # The rest of the members should be marked 0 if the are not -1.
                # for x,y in Email Address:
                # (x,y)
                'Leader': [1 if x == leader else y for (x, y) in zip(group_people['Email Address'], group_people['Leader'])],
                'Email Address': group_people['Email Address']
            }
            groups.append(pd.DataFrame(group_data))

            # print("leader")
            # print(df['Leader'])
            # df.style.highlight_max(subset=['Leader'] == 1, color='yellow')
            # # color all the leaders orange
            # print("bizza3")

            # def color_leader(col):
            #     if col.name == 'Leader':
            #         print("bizza4")
            #         return ['background-color: orange' if x == 1 else '' for x in col]

            # color all the leaders orange
            # data = pd.DataFrame(group_data)
            # data.style.apply(color_leader)

            # print("bizza5")
            # def color_leader(s):
            #     """
            #     function to color the leader
            #     """
            #     print("s")
            #     print(s['Leader'])
            #     if s['Leader'] == 1:
            #         return 'background-color: yellow'
            #     else:
            #         return ''

            # Add the used leader to the set for this Section
            used_leaders.add(leader)

    # Concatenate the groups list into a single DataFrame
    if groups:
        groups = pd.concat(groups)

        # Sort the groups DataFrame by Section, group, leader, and Email Address Address
        groups = groups.sort_values(
            by=['Section', 'Group', 'Leader', 'Email Address'], ascending=[True, True, False, True])

    else:
        print("No groups to concatenate.")

    # Reset the index of the 'groups' DataFrame to make it unique
    groups.reset_index(drop=True, inplace=True)

    # Reset the index of the 'groups' DataFrame to make it unique
    groups.reset_index(drop=True, inplace=True)

    # Define a function to color the 'Leader' column cells where the value is 1

    def color_leader(val):
        color = 'background-color: orange' if val == 1 else ''
        return color

    # Apply the color_leader function to the 'Leader' column in the 'groups' DataFrame
    styled_df = groups.style.map(color_leader, subset=['Leader'])

    pd.set_option('display.max_colwidth', 1000)
    # Save the styled DataFrame to an Excel file
    styled_df.to_excel('output.xlsx', engine='openpyxl', index=False)

    def download_excel():
        df.to_excel('downloaded_data.xlsx', index=False)
        with open('output.xlsx', 'rb') as f:
            data = f.read()
        return data

    excel_data = download_excel()
    st.download_button(label='Click to Download', data=excel_data,
                       file_name='cowl65_groups.xlsx', key='excel_download')

    # Display the styled DataFrame in the Streamlit app
    styled_df.set_properties(**{'max-width': '100px', 'font-size': '10pt'})

    styled_df.set_properties(subset=['Email Address'], **{'width': '500px'})
    st.markdown(styled_df.to_html(escape=False), unsafe_allow_html=True)
