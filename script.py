import pandas as pd
from itertools import cycle

# Read the Excel sheet into a pandas DataFrame
df = pd.read_excel('data/input.xlsx')

# Add an empty "Past Leaders" column to the DataFrame
# if there isn't one already
# This column will be used to track which leaders have already led a group
if 'Past Leaders' not in df.columns:
    df['Past Leaders'] = ''
else:
    df['Past Leaders'] = df['Past Leaders'].astype(
        str)  # Convert to string type

# Sort the DataFrame by time slot and first name
df.sort_values(['Time Slot'], inplace=True)

# Create a new column to track the leader status
df['Leader'] = 0

# Group the DataFrame by time slot
grouped = df.groupby('Time Slot')

# Create a cycle of unique leaders for each group
leaders = cycle(df['Email'])

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
        # Choose a new leader if they were a leader in the past. Ensure that each group will only have one new leader.
        leader = next(leaders)
        while leader in group_people['Past Leaders'].str.split(',', expand=True).values.flatten():
            leader = next(leaders)

        # Set the leader status for the assigned leader
        df.loc[df['Email'] == leader, 'Leader'] = 1

        # Update the past leaders column
        # Add the current leader to the past leaders column for each person in the group
        group_people.loc[:,
                         'Past Leaders'] = group_people['Past Leaders'] + leader + ','
        df.loc[df['Email'].isin(group_people['Email']),
               'Past Leaders'] = group_people['Past Leaders'].values

        # Append the group data to the groups list
        groups.append(pd.DataFrame({
            'Time Slot': time_slot,
            'Group': i + 1,
            'Leader': [1] + [0] * (3 + leftover),
            'Email': group_people['Email'],
            'Past Leaders': group_people['Past Leaders']
        }))

# Concatenate the groups list into a single DataFrame
groups = pd.concat(groups)

# Sort the groups DataFrame by time slot, group, and first name
groups.sort_values(by=['Time Slot', 'Group', 'Leader', 'Email', 'Past Leaders'], ascending=[
                   True, True, False, True, True], inplace=True)

# Write the groups DataFrame to a new Excel sheet
groups.to_excel('output.xlsx', index=False)
