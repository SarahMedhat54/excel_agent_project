from server import create_excel, create_chart

created = create_excel(
    columns=["Month", "Sales", "Profit"],
    rows=[
        ["Jan", 100, 20],
        ["Feb", 130, 30],
        ["Mar", 90, 15],
        ["Apr", 160, 45],
    ],
)
fid = created["file_id"]

print(create_chart(fid, "bar", "Month", ["Sales", "Profit"], title="Sales vs Profit"))
print(create_chart(fid, "line", "Month", ["Sales"], title="Sales Trend"))
print(create_chart(fid, "pie", "Month", ["Sales"], title="Sales Share"))
print("Open:", created["path"])