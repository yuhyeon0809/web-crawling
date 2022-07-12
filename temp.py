def csvtoexcel(filename_csv):
    r_csv = pd.read_csv(filename_csv)
    with pd.ExcelWriter("item_test.xlsx", mode="w", engine="openpyxl") as writer:
        r_csv.to_excel(writer, index = False) 
        writer.save() 