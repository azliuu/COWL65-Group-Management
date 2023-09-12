import pandas as pd
from itertools import cycle

# Read the Excel sheet into a pandas DataFrame
df = pd.read_excel('data/input.xlsx')

# # Add an empty "Past Leaders" column to the DataFrame
# df['Past Leaders'] = ''

# Sort the DataFrame by time slot and first name
print("this is timeslot", df['Time Slot'])
# print(df.sort_values(['Time Slot'], inplace=True))
df.sort_values(['Time Slot'], inplace=True)

# if leader column doesn't exist create leader column
if 'Leader' not in df.columns:
    df['Leader'] = 0

for index, row in df.iterrows():
    if (row['Leader'] == 1):
        df.at[index, 'Leader'] = -1

# Group the DataFrame by time slot
grouped = df.groupby('Time Slot')

# Create a new DataFrame to store the groups
groups = []

# Iterate over each time slot
for time_slot, time_slot_data in grouped:
    # Shuffle the time slot data to randomize the order
    time_slot_data = time_slot_data.sample(frac=1).reset_index(drop=True)

    # Calculate the number of groups needed
    num_groups = len(time_slot_data) // 4

    # Calculate how many people are leftover
    remainder = len(time_slot_data) % 4

    # Create a cycle of unique leaders for each group
    leaders = cycle(time_slot_data['Email'])

    # Create a set to track used leaders for this time slot
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
        leader = group_people.iloc[0]['Email']
        print("leader: ")
        print(leader)
        print("group people: ")
        print(group_people)
        # print(group_people[group_people['Leader']])
        # print("gp e")
        # print(group_people['Email'][0])
        # while (
        # leader in group_people['Email'].values or
        # leader in used_leaders or
        # (df.loc[df['Email'] == leader, 'Leader'].any() == -1)
        # ):
        # If the currently selected leader is marked as '-1',
        # try to find a member marked as '0' to be the new leader.
        # print("hihi")
        # print(df.loc[df['Email'] == leader, 'Leader'].any())
        # print("row")
        # print(df.loc[df['Email'] == leader])
        # print("leader")
        # # print(df.loc[df['Leader']])
        if df.loc[df['Email'] == leader, 'Leader'].any():
            new_leader_candidates = group_people[group_people['Leader'] == 0]
            if not new_leader_candidates.empty:
                leader = new_leader_candidates.iloc[0]['Email']
            else:
                # If no '0' leader candidates are available, select a member marked as '-1'.
                leader = group_people[group_people['Leader']
                                      == -1].iloc[0]['Email']
        else:
            # If the currently selected leader is not marked as '-1',
            # select a new leader from available members.
            leader = group_people[group_people['Leader']
                                  == 0].iloc[0]['Email']
            # print(leader)
        # Set the leader status for the assigned leader
        df.loc[df['Email'] == leader, 'Leader'] = 1
        # print(leaders)
        # print("this is leader: ")
        # print(leader)
        # print("this is value of leader: ")
        # print(df.loc[df['Email'] == leader, 'Leader'])

        # Update the past leaders column
        # df.loc[df['Email'].isin(group_people['Email']),
        #        'Past Leaders'] += leader + ','

        # Append the group data to the groups list, maintaining '-1' for leaders
        # print(group_people['Leader'])
        # print("this is email: ")
        # print(group_people['Email'])
        # print(group_people['Leader'])
        group_data = {
            'Time Slot': time_slot,
            'Group': i + 1,
            # choose the first member of the group as the leader if the leader is not -1.
            # The rest of the members should be marked 0 if the are not -1.
            # for x,y in email:
            # (x,y)
            'Leader': [1 if x == leader else y for (x, y) in zip(group_people['Email'], group_people['Leader'])],
            'Email': group_people['Email']
        }
        groups.append(pd.DataFrame(group_data))

        # Add the used leader to the set for this time slot
        used_leaders.add(leader)

# Concatenate the groups list into a single DataFrame
if groups:
    groups = pd.concat(groups)

    # Sort the groups DataFrame by time slot, group, leader, and email
    groups = groups.sort_values(
        by=['Time Slot', 'Group', 'Leader', 'Email'], ascending=[True, True, False, True])

    # Write the groups DataFrame to a new Excel sheet
    groups.to_excel('output.xlsx', index=False)
else:
    print("No groups to concatenate.")
