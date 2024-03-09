import dataframes as ds
for_room = []
data = []
sheet_names = list(ds.room_dict.keys())

for sheet_name, df in ds.room_dict.items():

    halltickets_by_branch = {}
    for branch in ds.bname.values():
        halltickets_by_branch[branch] = []
    for i, row in df.iterrows():
        for j, cell in row.iteritems():
            if cell != None:
                hallticket = cell.strip()
                if len(hallticket) != 10:
                    continue
                branch_code = hallticket[6:8]
                if branch_code not in ds.bname:
                    continue
                branch = ds.bname[branch_code]
                halltickets_by_branch[branch].append(hallticket)

    empty_branches = [branch for branch in halltickets_by_branch if len(halltickets_by_branch[branch]) == 0]
    for branch in empty_branches:
        del halltickets_by_branch[branch]

    for_room.append(halltickets_by_branch)

    collegecodes_by_branch = {}
    for branch, halltickets in halltickets_by_branch.items():
        collegecodes = []
        for hallticket in halltickets:
            collegecode = hallticket[2:4]
            if collegecode not in collegecodes:
                collegecodes.append(collegecode)
        collegecodes_by_branch[branch] = collegecodes

    keynames = list(halltickets_by_branch.keys())
    data.append({'keynames':keynames, 'collegecodes_by_branch': collegecodes_by_branch})
    print(data)
    code=[]
    for i in range(0,len(data)):
        code.append(data[i]['collegecodes_by_branch'].values())
for d in for_room:
    for key in d:
        print(f"{key}: {len(d[key])}")
output_dict = {}

for i in range(len(sheet_names)):
    halltickets_by_branch = for_room[i]

    sheet_dict = {}
    for branch in halltickets_by_branch:
        count = len(halltickets_by_branch[branch])
        sheet_dict[branch] = count

    output_dict[sheet_names[i]] = sheet_dict

print(output_dict)
merge_dict={}
for key, value in output_dict.items():
    count = len(value.keys())
    merge_dict[key] = count
print(merge_dict)