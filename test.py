from prettytable import PrettyTable
    
x = PrettyTable()

x.field_names = ["City name", "Area", "Population", "Annual Rainfall"]

x.add_row(["Adelaide", 1295, 1158259])


print(x)