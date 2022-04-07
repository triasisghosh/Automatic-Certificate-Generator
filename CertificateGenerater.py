import csv
import os.path
from os import listdir, system
from tkinter import *
from tkinter import colorchooser, filedialog, messagebox
from tkinter.ttk import Combobox, Scrollbar

try:
    import cv2
except ModuleNotFoundError or ImportError:
    system('python -m pip install opencv-python')
    import cv2

try:
    from tkcalendar import Calendar
except ModuleNotFoundError or ImportError:
    system('python -m pip install tkcalendar')
    from tkcalendar import Calendar

try:
    from PIL import ImageTk, Image, ImageDraw, ImageFont
except ModuleNotFoundError or ImportError:
    system('python -m pip install pillow')
    from PIL import ImageTk, Image, ImageDraw, ImageFont


class MainBox:
    root = Tk()
    root.title('Certificate Generator')
    root.geometry('600x600')
    other_fields = []
    csv_fields = []
    image_fields = []
    i = 0

    def __init__(self):
        canv_frame = Frame(self.root, height=400)
        canv_frame.pack(fill=BOTH, expand=1)
        base_canvas = Canvas(canv_frame)
        base_canvas.pack(side=LEFT, fill=BOTH, expand=1)

        scroll_bar = Scrollbar(canv_frame, orient=VERTICAL, command=base_canvas.yview)
        scroll_bar.pack(side=RIGHT, fill=Y)

        def onmousewheel(event):
            base_canvas.yview_scroll(-1 * (event.delta // 120), "units")

        base_canvas.configure(yscrollcommand=scroll_bar.set)
        base_canvas.bind('<Configure>',
                         lambda e: base_canvas.configure(scrollregion=base_canvas.bbox('all')))
        base_canvas.bind_all('<MouseWheel>', onmousewheel)
        self.main = Frame(base_canvas)
        base_canvas.create_window((0, 0), window=self.main, anchor=NW)

        execute_button_frame = Frame(self.root)
        execute_button_frame.pack()
        self.preview_button = Button(execute_button_frame, text='PREVIEW', bg='Blue', fg='white', relief='groove',
                                     font=('Ariel', 15))
        self.preview_button.grid(row=0, column=0, pady=5, padx=10)

        self.generate_button = Button(execute_button_frame, text='GENERATE', bg='Red', fg='white', relief='groove',
                                      font=('Ariel', 15))
        self.generate_button.grid(row=0, column=1)

        file_selection = Frame(self.main)
        file_selection.pack()
        self.select_csv = self.DirctField(0, 'Select CSV File', file_selection, self.pick_csv)
        self.select_template = self.DirctField(1, 'Select Template', file_selection,
                                               lambda: self.pick_template(self.select_template.entry))
        self.select_fold = self.DirctField(2, 'Select Destination Folder', file_selection, self.pick_destination_folder)

        ot_add_frame = Frame(self.main)
        ot_add_frame.pack(padx=5, pady=5)

        def no_keyboard_input(input):
            return False

        nki = ot_add_frame.register(no_keyboard_input)

        def int_only(input):
            if input.isdigit() or input == '':
                return True
            else:
                return False

        verify_int = ot_add_frame.register(int_only)

        Label(ot_add_frame, text='Start row:', fg='black', font=('Ariel', 10)).grid(row=0, column=0, pady=3)
        self.start_row = Combobox(ot_add_frame, width=3, values=[i for i in range(10)])
        self.start_row.current(0)
        self.start_row.config(validate='key', validatecommand=(verify_int,'%P'))
        self.start_row.grid(row=0, column=1, pady=3)

        Label(ot_add_frame, text='Key column:', fg='black', font=('Ariel', 10)).grid(row=0, column=2, pady=3)
        self.key_column = Combobox(ot_add_frame, width=3, values=[i for i in range(10)])
        self.key_column.current(0)
        self.key_column.config(validate='key', validatecommand=(verify_int,'%P'))
        self.key_column.grid(row=0, column=3, pady=3)

        add_other_button = Button(ot_add_frame, text='Add Field', command=self.add_other)
        add_other_button.grid(row=1, column=0)
        self.add_other_entrytype = Combobox(ot_add_frame, values=['From CSV file', 'Constant', 'Image', 'Date'],
                                            width=10)
        self.add_other_entrytype.config(validate="key", validatecommand=(nki, '%P'))
        self.add_other_entrytype.current(1)
        self.add_other_entrytype.grid(row=1, column=1)

        delete_other_button = Button(ot_add_frame, text='Delete Field', command=self.del_other)
        delete_other_button.grid(row=1, column=2)
        self.delete_other_entrytype = Combobox(ot_add_frame, width=10)
        self.delete_other_entrytype.config(validate="key", validatecommand=(nki, '%P'))
        self.delete_other_entrytype.grid(row=1, column=3)

        self.Name = self.DataField('Name Properties', 'From CSV file', self.main, self.select_template.entry)
        self.Date_field = self.DataField('Date Properties', 'Date', self.main, self.select_template.entry)

    class DirctField:
        def __init__(self, r, lbl, win, operation):
            Label(win, text=lbl, fg='black', font=('Ariel', 10)).grid(row=r, column=0, padx=5, pady=5)
            self.entry = Entry(win, width=50)
            self.entry.grid(row=r, column=1)
            button = Button(win, text='Select File', command=operation)
            button.grid(row=r, column=2)

    class DataField:
        def __init__(self, label, entry_type, main_root, template):
            self.root_frame = Frame(main_root, pady=20)
            self.root_frame.pack()
            self.verify_integer = self.root_frame.register(self.callback)
            lable_frame = Frame(self.root_frame)
            lable_frame.pack()
            self.Source = entry_type
            self.Title = label
            self.Required = BooleanVar(value=True)
            Label(lable_frame, text=label, fg='blue', font=('Ariel', 14)).grid(row=0, column=0)
            Checkbutton(lable_frame, text='Required', onvalue=True, offvalue=False, variable=self.Required).grid(row=0,
                                                                                                                 column=1,
                                                                                                                 padx=5)

            def no_keyboard_input(input):
                return False

            nki = self.root_frame.register(no_keyboard_input)

            if entry_type == 'From CSV file':
                win1 = Frame(self.root_frame)
                win1.pack()
                Label(win1, text='Index of column:', fg='black', font=('Ariel', 10)).grid(column=0, padx=5, pady=5)
                self.entry1 = Combobox(win1, values=[i for i in range(10)], width=3)
                self.entry1.config(validate="key", validatecommand=(self.verify_integer, '%P'))
                self.entry1.grid(row=0, column=1)
            elif entry_type == 'Constant':
                win1 = Frame(self.root_frame)
                win1.pack()
                Label(win1, text='Constant text', fg='black', font=('Ariel', 10)).grid(column=0, padx=5, pady=5)
                self.entry1 = Entry(win1, width=20)
                self.entry1.grid(row=0, column=1)
            elif entry_type == 'Image':
                win1 = Frame(self.root_frame)
                win1.pack()
                Label(win1, text='Image', fg='black', font=('Ariel', 10)).grid(column=0, padx=5, pady=5)
                self.entry1 = Entry(win1, width=20)
                self.entry1.grid(row=0, column=1)
                button1 = Button(win1, text='Select File', command=lambda: MainBox.pick_template(self, self.entry1))
                button1.grid(row=0, column=2)
            elif entry_type == 'Date':
                win1 = Frame(self.root_frame)
                win1.pack()
                Label(win1, text='Date', fg='black', font=('Ariel', 10)).grid(column=0, padx=5, pady=5)
                self.entry1 = Entry(win1, width=12)
                self.entry1.config(validate='key', validatecommand=(nki, '%P'))
                self.entry1.grid(row=0, column=1)
                Button(win1, text='Select Date', command=lambda: self.pickdate(self.entry1)).grid(row=0, column=2)
            if self.Source != 'Image':
                win2 = Frame(self.root_frame)
                win3 = Frame(self.root_frame)
                win4 = Frame(self.root_frame)

                win2.pack()
                win3.pack()
                win4.pack()

                Label(win2, text='Font style', fg='black', font=('Ariel', 10)).grid(row=0, column=0, padx=5, pady=5)
                self.entry2 = Combobox(win2, values=[i for i in self.fonts()], width=10)
                self.entry2.config(validate="key", validatecommand=(nki, '%P'))
                self.entry2.grid(row=0, column=1)
                try:
                    prvw_img_main = Image.open('preview.png')
                    prvw_img_show = ImageTk.PhotoImage(prvw_img_main)
                    self.canvas1 = Label(win2, image=prvw_img_show)
                    self.canvas1.image = prvw_img_show
                    self.entry2.bind("<<ComboboxSelected>>", self.preview_on_canvas)
                    self.canvas1.grid(row=0, column=3)
                except FileNotFoundError:
                    pass

                Label(win3, text='Font size', fg='black', font=('Ariel', 10)).grid(row=0, column=0, padx=5, pady=5)
                self.entry3 = Combobox(win3, values=[i for i in range(0, 101, 4)], width=5)
                self.entry3.config(validate="key", validatecommand=(self.verify_integer, '%P'))
                self.entry3.grid(row=0, column=1)

                Label(win4, text='Colour code', fg='black', font=('Ariel', 10)).grid(row=0, column=0, padx=5, pady=5)

                self.entry4_i = Entry(win4, width=5, state='disabled')
                self.entry4_ii = Entry(win4, width=5, state='disabled')
                self.entry4_iii = Entry(win4, width=5, state='disabled')

                self.entry4_i.grid(row=0, column=1)
                Label(win4, text='x', fg='black', font=('Ariel', 10)).grid(row=0, column=2)
                self.entry4_ii.grid(row=0, column=3)
                Label(win4, text='x', fg='black', font=('Ariel', 10)).grid(row=0, column=4)
                self.entry4_iii.grid(row=0, column=5)

                self.button2 = Button(win4, text='       ', relief='ridge',
                                      command=lambda: self.pick_colour(self.entry4_i, self.entry4_ii, self.entry4_iii))
                self.button2.grid(row=0, column=6, padx=5)
                Label(win4, text='Select Color', fg='black', font=('Ariel', 10)).grid(row=0, column=7)

            win5 = Frame(self.root_frame)
            win5.pack()
            Label(win5, text='Find position on template', fg='black', font=('Ariel', 10)).grid(row=0, column=0, padx=5,
                                                                                               pady=5)
            self.entry5_i = Entry(win5, width=5)
            self.entry5_ii = Entry(win5, width=5)
            self.entry5_i.config(validate="key", validatecommand=(self.verify_integer, '%P'))
            self.entry5_ii.config(validate="key", validatecommand=(self.verify_integer, '%P'))

            self.entry5_i.grid(row=0, column=1)
            Label(win5, text='x', fg='black', font=('Ariel', 10)).grid(row=0, column=2)
            self.entry5_ii.grid(row=0, column=3)

            button3 = Button(win5, text='Choose Position',
                             command=lambda: self.position_finder(self.entry5_i, self.entry5_ii, template))
            button3.grid(row=0, column=4, padx=5)

        def pickdate(self, entry):
            date_window = Tk()
            date_window.title('Date Picker')
            date_window.geometry('500x300')

            def callback(input):
                return False

            nki = date_window.register(callback)

            dfmc_win = Frame(date_window)
            dfmc_win.pack()
            df_win = Frame(date_window)
            df_win.pack()
            Label(dfmc_win, text='Customized date format:', fg='black', font=('Ariel', 10)).grid(row=0, column=0,
                                                                                                 padx=5, pady=5)
            date_format_customized = []
            for i in range(4):
                if i != 3:
                    date_format_customized.append(Combobox(dfmc_win, width=5,
                                                           values=['d', 'dd', 'm', 'mm', 'yy', 'yyyy']))
                    date_format_customized[-1].current(2 * i + 1)
                    date_format_customized[-1].config(validate="key", validatecommand=(nki, '%P'))
                    date_format_customized[-1].grid(row=0, column=i + 1, padx=5)
                else:
                    Label(dfmc_win, text='Delimiter', fg='black', font=('Ariel', 10)).grid(row=0, column=4, padx=5)
                    date_format_customized.append(Combobox(dfmc_win, width=1, values=['.', '-', '/', '|', ',']))
                    date_format_customized[-1].current(1)
                    date_format_customized[-1].config(validate="key", validatecommand=(nki, '%P'))
                    date_format_customized[-1].grid(row=0, column=5, padx=5)
            date_format = Combobox(df_win, width=12,
                                   values=['dd-mm-yyyy', 'dd-mm-yy', 'dd.mm.yyyy', 'dd.mm.yy', 'dd/mm/yyyy',
                                           'dd/mm/yy'])
            date_format.current(0)
            date_format.config(validate="key", validatecommand=(nki, '%P'))

            def select_date_format():
                s = date_format_customized[-1].get().join(list(j.get() for j in date_format_customized[:3]))
                date_format.config(validate='none')
                date_format.delete(0, END)
                date_format.insert(END, string=s)
                date_format.config(validate='key')

            Button(dfmc_win, text='Select', command=select_date_format).grid(row=0, column=6)

            Label(df_win, text='Date format:', fg='black', font=('Ariel', 10)).grid(row=0, column=0, padx=5, pady=5)
            date_format.grid(row=0, column=1, padx=5, pady=5)
            cal = Calendar(date_window, selectmode='day')
            cal.pack()

            def grab_date():
                try:
                    cal.config(date_pattern=date_format.get())
                    s = cal.get_date()
                    entry.config(validate='none')
                    entry.delete(0, END)
                    entry.insert(END, s)
                    entry.config(validate='key')
                except ValueError:
                    messagebox.showerror(title='Date error', message='Invalid date format.')

            Button(date_window, text='Select Date', command=grab_date).pack(pady=5)
            date_window.mainloop()

        def pick_colour(self, a, b, c):
            chosen_colour = colorchooser.askcolor()
            a.configure(state='normal')
            b.configure(state='normal')
            c.configure(state='normal')
            if chosen_colour[1] is not None:
                a.delete(0, 'end')
                a.insert(0, chosen_colour[0][0])

                b.delete(0, 'end')
                b.insert(0, chosen_colour[0][1])

                c.delete(0, 'end')
                c.insert(0, chosen_colour[0][2])

                self.button2.config(bg='#%02x%02x%02x' % (int(a.get()), int(b.get()), int(c.get())))
            a.configure(state='disabled')
            b.configure(state='disabled')
            c.configure(state='disabled')

        def position_finder(self, a, b, t):
            def getpoint(event, x, y, flags, params):
                if event == cv2.EVENT_LBUTTONDOWN:
                    a.delete(0, 'end')
                    a.insert(0, str(x))
                    b.delete(0, 'end')
                    b.insert(0, str(y))

            pic = t.get()
            try:
                img = open(pic, 'r')
            except OSError or FileNotFoundError:
                messagebox.showerror('Error', 'Template path is empty or invalid.')
                return
            cv2.namedWindow('Template', cv2.WINDOW_NORMAL)
            img = cv2.imread(pic)
            size = img.shape[:2]
            cv2.line(img, (0, size[0] // 2), (size[1], size[0] // 2), (0, 255, 0), 2)
            cv2.line(img, (size[1] // 2, 0), (size[1] // 2, size[0]), (0, 255, 0), 2)
            cv2.imshow('Template', img)
            cv2.setMouseCallback('Template', getpoint)

        def fonts(self):
            font_address = 'C:\\Windows\\Fonts'
            try:
                font_list = listdir(font_address)
            except OSError or FileNotFoundError:
                return []
            font_list = [i for i in font_list if any(x in i for x in ['.TTF', '.ttf', '.OTF', '.otf'])]
            return font_list

        def preview_on_canvas(self, event):
            try:
                img = Image.open('preview.png')
            except OSError or FileNotFoundError:
                return
            text = self.entry2.get()
            if text == '':
                return
            try:
                fnt = ImageFont.truetype('C:\\Windows\\Fonts\\' + text, 23)
            except OSError or FileNotFoundError:
                messagebox.showerror(title='Error', message='Font not found.')
                return
            draw = ImageDraw.Draw(img)
            draw.text((2, 1), 'Sample', fill=(0, 0, 0), font=fnt)
            img = ImageTk.PhotoImage(img)
            self.canvas1.config(image=img)
            self.canvas1.image = img

        def callback(self, input):
            if input.isdigit() or input == '':
                return True
            else:
                return False

    def pick_csv(self):
        data_list = filedialog.askopenfilename(filetype=[("CSV file", "*.csv")], title='Select Name List')
        self.select_csv.entry.delete(0, 'end')
        self.select_csv.entry.insert(0, data_list)

    def pick_template(self, path_field):
        img_path = filedialog.askopenfilename(
            filetype=[("PNG file", "*.png"), ("JPG file", "*.jpg"), ("JPEG file", "*.jpeg")], title='Select Template')
        path_field.delete(0, 'end')
        path_field.insert(0, img_path)

    def pick_destination_folder(self):
        dest = filedialog.askdirectory(title='Select Destination Folder')
        self.select_fold.entry.delete(0, 'end')
        self.select_fold.entry.insert(0, dest)

    def add_other(self):
        s = self.add_other_entrytype.get()
        if s != '':
            self.i = self.i + 1
            self.other_fields.append(self.DataField(f'Other {self.i}', s, self.main, self.select_template.entry))
            if s == 'From CSV file':
                self.csv_fields.append(self.other_fields[-1])
                print(len(self.csv_fields))
            elif s == 'Image':
                self.image_fields.append(self.other_fields[-1])
            self.delete_other_entrytype.config(values=[i.Title for i in self.other_fields])
        else:
            return

    def del_other(self):
        s = self.delete_other_entrytype.get()
        if s != '':
            for i in range(len(self.other_fields)):
                if self.other_fields[i].Title == s:
                    if self.other_fields[i].Source == 'From CSV file':
                        for j in range(len(self.csv_fields)):
                            if self.csv_fields[i].Title == s:
                                self.csv_fields.pop(i)
                    elif self.other_fields[i].Source == 'Image':
                        for j in range(len(self.image_fields)):
                            if self.image_fields[i].Title == s:
                                self.image_fields.pop(i)
                    self.other_fields[i].root_frame.destroy()
                    del self.other_fields[i]
                    self.delete_other_entrytype.config(values=[i.Title for i in self.other_fields])
                    break

    def csv_indices_check(self, n):
        if any(int(i.entry1.get()) < 0 or int(i.entry1.get()) > n for i in self.csv_fields):
            return False
        else:
            return True

    def image_exist(self):
        if all(file_check(i.entry1.get(), 'Insert Image', ['.png', '.jpg', '.jpeg']) for i in self.image_fields):
            # for i in self.image_fields:
            # if file_check(i.entry1.get()) and any(i.endswith(j)for j in ['.png','.jpg','.jpeg'])
            return True
        else:
            messagebox.showerror(title='Error', message='Insert image field image path invalid.')
            return False

    def start_tuple_check(self, n):
        if int(self.start_row.get()) > n:
            messagebox.showerror(title='Error', message='Start row out of range.')
            return False
        else:
            return True


gui = MainBox()


class FieldDataSet:
    insertion_dataset = []

    def __init__(self, mainbox):
        if mainbox.Name.Required.get():
            self.insertion_dataset.append(self.FieldDataBlock(mainbox.Name))
        if mainbox.Date_field.Required.get():
            self.insertion_dataset.append(self.FieldDataBlock(mainbox.Date_field))
        for i in mainbox.other_fields:
            if i.Required.get():
                self.insertion_dataset.append(self.FieldDataBlock(i))

    class FieldDataBlock:
        def __init__(self, record):
            if record.Source == 'Constant' or record.Source == 'Date':
                self.source = 'constant'

            elif record.Source == 'From CSV file':
                self.source = 'csv'
            elif record.Source == 'Image':
                self.source = 'image'
            self.info = record.entry1.get()
            if self.source != 'image':
                self.font = 'C:\\Windows\\Fonts\\' + record.entry2.get()
                self.size = int(record.entry3.get())
                self.color = (int(record.entry4_i.get()), int(record.entry4_ii.get()), int(record.entry4_iii.get()))
            self.position = (int(record.entry5_i.get()), int(record.entry5_ii.get()))


def image_centering(xy, imgsize):
    xy = list(xy)
    len_i = imgsize
    xy[0] = xy[0] - len_i[0] // 2
    xy[1] = xy[1] - len_i[1] // 2
    return tuple(xy)


def text_centering(xy, fnt, string):
    xy = list(xy)
    len_s = fnt.getsize(string)
    xy[0] = xy[0] - len_s[0] // 2
    xy[1] = xy[1] - len_s[1] // 2
    return tuple(xy)


def certificate(tupl, datalist, template_path, destination_path='', savekeyindex=0, s=True):
    basepic = Image.open(template_path)
    for i in datalist:
        if i.source != 'image':
            writable_image = ImageDraw.Draw(basepic)
            if i.source == 'csv':
                try:
                    tm = tupl[int(i.info)]
                except IndexError:
                    messagebox.showerror(title='Error', message='CSV column not found.')
                    tm = ''
            else:
                tm = i.info
            try:
                fnt = ImageFont.truetype(i.font, i.size)
            except OSError or FileNotFoundError:
                messagebox.showerror(title='Error', message='Font not found.')
                return
            fc = i.color
            fp = text_centering(i.position, fnt, tm)
            writable_image.text(xy=fp, text=tm, fill=fc, font=fnt)
        else:
            tm = Image.open(i.info)
            fp = image_centering(i.position, tm.size)
            basepic.paste(tm, fp)
    basepic.show()
    if s:
        try:
            basepic.save(destination_path + '\\' + tupl[savekeyindex]+'.png', format('png'))
        except IndexError or tupl[savekeyindex] == '':
            messagebox.showerror(title='Error', message='Key value missing or empty. Couldn\'t save')
    basepic.close()
    return


def empty_check(field_list):
    for i in field_list:
        if i.Required.get():
            if i.Source != 'Image':
                if any(x == '' for x in
                       [i.entry1.get(), i.entry2.get(), i.entry3.get(), i.entry4_i.get(), i.entry4_ii.get(),
                        i.entry4_iii.get(), i.entry5_i.get(), i.entry5_ii.get()]):
                    messagebox.showerror(title='Error', message='All fields are mandatory. Delete if not needed.')
                    return False
            else:
                if any(x == '' for x in [i.entry1.get(), i.entry5_i.get(), i.entry5_ii.get()]):
                    messagebox.showerror(title='Error', message='All fields are mandatory. Delete if not needed.')
                    return False
        else:
            return True
    return True


def file_check(path, name, extensions=[]):
    if os.path.isfile(path):
        if any(path.lower().endswith(x) for x in extensions):
            return True
        else:
            messagebox.showerror(title='Error', message='Select' + ', '.join(extensions) + 'file formats only.')
    else:
        messagebox.showerror(title='Error', message=name + " file path doesn't exist.")
        return False


def folder_check(path):
    if os.path.isdir(path):
        return True
    else:
        messagebox.showerror(title='Error', message="Folder path doesn't exist.")
        return False


def generate(preview=False):
    template_path = gui.select_template.entry.get()
    csv_path = gui.select_csv.entry.get()
    destination_path = gui.select_fold.entry.get()

    if file_check(template_path, 'Template', ['.jpg', '.jpeg', '.png']) and file_check(csv_path, 'CSV',
                                                                                       ['.csv']) and gui.image_exist():
        if folder_check(destination_path):
            cf = open(csv_path, 'r')
            csv_file = list(csv.reader(cf))
            if not gui.start_tuple_check(len(csv_file)):
                return
        else:
            return
    else:
        return
    if empty_check([gui.Name, gui.Date_field]) and empty_check(gui.other_fields):
        start_tuple_index = int(gui.start_row.get())
        key_column = int(gui.key_column.get())
        data_object = FieldDataSet(gui)
    else:
        return

    if preview:
        certificate(csv_file[start_tuple_index], data_object.insertion_dataset, template_path, s=False)
    else:
        for tupl in csv_file[start_tuple_index:]:
            certificate(tupl, data_object.insertion_dataset, template_path, destination_path,
                        savekeyindex=key_column, s=True)
    del data_object.insertion_dataset[:]
    del data_object
    return


gui.preview_button.configure(command=lambda: generate(preview=True))
gui.generate_button.configure(command=generate)
gui.root.mainloop()