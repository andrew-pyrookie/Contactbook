import xlwt
import xlrd

workbook = xlwt.Workbook()
contacts = workbook.add_sheet("Contacts")
contacts.write(0, 2, "Email Addr")
contacts.write(0, 3, "City")
contacts_data = {}


def save_workbook():
    workbook.save("Contact.xls")


def add_contact(contacts):
    Name = input("Name: ")
    number = input("Contact: ")
    contacts_data[Name] = number

    # Find the next available row in the sheet
    row = len(contacts_data)

    # Write the contact information to the sheet
    contacts.write(row, 0, Name)
    contacts.write(row, 1, number)

    save_workbook()





def delete(contacts):
    name = input("Which name_contact to delete: ")
    del contacts[name]
    save_workbook()


def search(contacts_data):
    name = input("Which contact name to search: ")
    items = contacts_data.items()
    for name, contacts in contacts_data.items():
        print(name, contacts)


def view():
    try:
        workbook = xlrd.open_workbook("Contact.xls")
        sheet = workbook.sheet_by_index(0)
        for row in range(sheet.nrows):
            for col in range(sheet.ncols):
                cell_value = sheet.cell_value(row, col)
                print(cell_value, end="\t")
            print()
    except xlrd.XLRDError:
        print("Unable to open 'Contact.xls' file.")


while True:
    print("Welcome To managing Your ContactBook")
    print("")
    print('''
      To save 'S':
      To Add 'A':
      To delete 'D':
      Delete All 'C':
      Search Contact 'M':
      To View All Cont'V':
      ''')

    print('')
    action = input("Your Action (S, A, D, C): ")
    print('')

    if action.upper() == 'S':
        save_workbook()
        break
    elif action.upper() == 'A':
        add_contact(contacts)
        break
    elif action.upper() == 'D':
        delete(contacts)
        break
    elif action.upper() == 'C':
        contacts_data.clear()
        break
    elif action.upper() == 'M':
        search(contacts_data)
        break
    elif action.upper() == 'V':
        view()
        break
