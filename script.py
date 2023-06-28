import pandas as pd
from itertools import cycle

# Read the Excel sheet into a pandas DataFrame
df = pd.read_excel('data/input.xlsx')

# Add an empty "Past Leaders" column to the DataFrame
df['Past Leaders'] = ''

# Sort the DataFrame by time slot and first name
df.sort_values(['time slot', 'First Name'], inplace=True)

# Create a new column to track the leader status
df['Leader'] = 0

# Group the DataFrame by time slot
grouped = df.groupby('time slot')

# Create a cycle of unique leaders for each group
leaders = cycle(df['First Name'].unique())

# Create a new DataFrame to store the groups
groups = []

# Iterate over each time slot
for time_slot, time_slot_data in grouped:
    # Shuffle the time slot data to randomize the order
    time_slot_data = time_slot_data.sample(frac=1).reset_index(drop=True)

    # Calculate the number of groups needed
    num_groups = len(time_slot_data) // 4

    # Iterate over each group
    for i in range(num_groups):
        # Get the people for the current group
        group_people = time_slot_data.iloc[i * 4:(i + 1) * 4]

        # Find a unique leader for the group
        leader = next(leaders)
        while leader in time_slot_data['Past Leaders'].values or leader in group_people['First Name'].values:
            leader = next(leaders)

        # Set the leader status for the assigned leader
        df.loc[df['First Name'] == leader, 'Leader'] = 1

        # Update the past leaders column
        df.loc[df['First Name'].isin(
            group_people['First Name']), 'Past Leaders'] += leader + ','

        # Append the group data to the groups list
        groups.append(pd.DataFrame({
            'Group': i + 1,
            'time slot': time_slot,
            'First Name': group_people['First Name'],
            'Leader': [1] + [0] * 3
        }))

# Concatenate the groups list into a single DataFrame
groups = pd.concat(groups)

# Sort the groups DataFrame by time slot, group, and first name
groups.sort_values(['time slot', 'Group', 'First Name'], inplace=True)

# Write the groups DataFrame to a new Excel sheet
groups.to_excel('output.xlsx', index=False)
