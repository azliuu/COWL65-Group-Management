import pandas as pd
from itertools import cycle

# Read the Excel sheet into a pandas DataFrame
df = pd.read_excel('data/input.xlsx')

# Add an empty "Past Leaders" column to the DataFrame
df['Past Leaders'] = ''

# Sort the DataFrame by time slot and first name
df.sort_values(['time slot'], inplace=True)

# Create a new column to track the leader status
df['Leader'] = 0

# Group the DataFrame by time slot
grouped = df.groupby('time slot')

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

    # Calculate how many people leftover
    remainder = len(time_slot_data) % 4

    # Iterate over each group
    for i in range(num_groups):
        # Get the people for the current group
        group_people = time_slot_data.iloc[i * 4:(i + 1) * 4]

        # Check if a leftover person needs to be added to a group
        leftover = 0
        if i < remainder:
            leftover_person = time_slot_data.iloc[len(time_slot_data) - i - 1: len(time_slot_data) - i]
            group_people = group_people._append(leftover_person)
            leftover = 1

        # Find a unique leader for the group
        leader = next(leaders)
        while leader in time_slot_data['Past Leaders'].values or leader not in group_people['Email'].values:
            leader = next(leaders)

        # Set the leader status for the assigned leader
        df.loc[df['Email'] == leader, 'Leader'] = 1

        # Update the past leaders column
        df.loc[df['Email'].isin(
            group_people['Email']), 'Past Leaders'] += leader + ','

        # Append the group data to the groups list
        groups.append(pd.DataFrame({
            'Time Slot': time_slot,
            'Group': i + 1,
            'Leader': [1] + [0] * (3 + leftover),
            'Email': group_people['Email']
        }))

# Concatenate the groups list into a single DataFrame
groups = pd.concat(groups)

# Sort the groups DataFrame by time slot, group, and first name
groups.sort_values(by=['Time Slot', 'Group', 'Leader', 'Email'], ascending=[True, True, False, True], inplace=True)

# Write the groups DataFrame to a new Excel sheet
groups.to_excel('output.xlsx', index=False)
