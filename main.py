from tkinter import *
from tkinter.filedialog import askopenfile
import os
from pandas import read_excel
import csv

"""this is just a check"""

class Check:
    """this class contains methods that returns booleans
    to notify if entry formats are followed
    """

    def __init__(self, data_input):
        """assign values to attributes

        Args:
            data_input (excel file as array): excel array from pandas
        """
        self.input = data_input
        self.nrow = len(data_input)
        self.ncol = len(data_input.columns)

        self.dict = {}
        for i in range(self.nrow):
            self.dict[i] = data_input.iloc[i].to_list()
        self.nlist = len(self.dict)

        self.length_True = {
            0: [10, 0],
            1: [18, 1],
            2: [20, 0],
            3: [3,  0],
            4: [34, 1],
            5: [18, 0],
            6: [1,  2],
            7: [18, 1],
            8: [18, 0],
            9: [7,  0],
            10: [30, 1]
        }
        self.length_False = {
            0: [10, 0],
            1: [18, 1],
            2: [20, 0],
            3: [3,  0],
            4: [34, 1],
            5: [18, 0],
            6: [7,  0],
            7: [40, 1]
        }

    def simple_check(self):
        """checks if input adhered to general format

        Returns:
            boolean: True if pass, False if else
        """
        for i in range(self.nrow):

            # Check for differences in format and size of input
            sl_length = len(self.dict[i])    # specific list length

            if self.nlist != self.nrow:
                return False
            if sl_length not in (8, 11):
                return False

        return True

    def check_type(self):
        """check type of conversion and if any values over limit
        """
        def extract(row, column):
            """extract specific value as string through specific inputs

            Args:
                row (int): row number
                column (int): column number

            Returns:
                string: specific value as string
            """
            return str(self.dict[row][column]).replace(".", " ")

        def check_strtype(input_string, str_type):
            """checks if inputed string follows type

            Args:
                input_string (str): inputed string
                str_type (int): number that references type of string

            Returns:
                boolean: True if string type followed, otherwise False
            """
            i_str = input_string

            str_test = i_str.replace(" ", "")
            if str_type == 0:
                return str_test.isnumeric()
            elif str_type == 1:
                return str_test.isalnum()
            else:
                return str_test.isalpha()

        c_type = False
        format_ad = True

        if self.ncol == 11:
            c_type = True

        for r in range(self.nrow):

            if c_type == True:
                for c in range(11):
                    object_val = extract(r, c)
                    object_len = len(object_val)
                    max_len = self.length_True[c][0]
                    o_type = self.length_True[c][1]

                    if object_len > max_len:
                        format_ad = False
                    if check_strtype(object_val, o_type) == False:
                        format_ad = False

                    if format_ad == False:
                        break

            else:
                for c in range(8):
                    object_val = extract(r, c)
                    object_len = len(object_val)
                    max_len = self.length_False[c][0]
                    o_type = self.length_False[c][1]

                    if object_len > max_len:
                        format_ad = False
                    if check_strtype(object_val, o_type) == False:
                        format_ad = False

                    if format_ad == False:
                        break

        return(c_type, format_ad)


class Format:
    """this class contains the two methods that help format inputed file
    """

    def __init__(self, input_dictionary, input_type, number_row, number_column):
        """assign values to attributes

        Args:
            input_dictionary (dictionary): inputed file as dictionary
            input_type (boolean): boolean checking if "with Comprobante Fiscal"
            number_row (int): total number of relevant rows in inputed file
            number_column (int): total number of relevant columns in inputed file
        """
        self.dict = input_dictionary
        self.itype = input_type
        self.nrow = number_row
        self.ncol = number_column

        self.ilist = []
        for i in range(self.nrow):
            self.ilist.append(input_dictionary.iloc[i].to_list())
        self.flist = []

    def totxt(self):
        """modify self.flist to ideal list for output in txt format

        Returns:
            list: contains strings that follow format
        """
        # functions that modify each object to conform to format
        def cuenta_o(i_object):
            i_string = str(i_object)
            diff = 10 - len(i_string)
            return ('0' * diff) + i_string

        def rfc_o(i_object):
            i_string = str(i_object)
            diff = 18 - len(i_string)
            return i_string + (" " * diff)

        def clabe(i_object):
            i_string = str(i_object)
            diff = 20 - len(i_string)
            return ('0' * diff) + i_string

        def banco_b(i_object):
            i_string = str(i_object)
            diff = 3 - len(i_string)
            return ('0' * diff) + i_string

        def nombre_b(i_object):
            i_string = str(i_object)
            diff = 34 - len(i_string)
            return i_string + (" " * diff)

        def monto(i_object):
            i_string = "{:.2f}".format(i_object).replace(".", "")
            diff = 18 - len(i_string)
            return ('0' * diff) + i_string

        def c_fiscal(i_object, con_cf):
            i_string = str(i_object)
            if con_cf == True:
                if i_string == 'S':
                    return i_string
                else:
                    return 'S'
            else:
                pass

        def rfc_b(i_object, con_cf):
            i_string = str(i_object)
            if con_cf == True:
                diff = 18 - len(i_string)
                return i_string + (" " * diff)
            else:
                pass

        def iva(i_object, con_cf):
            i_string = "{:.2f}".format(i_object).replace(".", "")
            if con_cf == True:
                diff = 18 - len(i_string)
                return ('0' * diff) + i_string
            else:
                pass

        def ref_n(i_object):
            i_string = str(i_object)
            diff = 7 - len(i_string)
            return ('0' * diff) + i_string

        def c_pago(i_object, con_cf):
            i_string = str(i_object)
            max_lim = 40
            if con_cf == True:
                max_lim = 30
            diff = max_lim - len(i_string)
            return i_string + (" " * diff)

        for r in range(self.nrow):
            current_l = self.ilist[r]
            with_cf = bool(self.itype)

            first_string = (
                cuenta_o(current_l[0])
                + rfc_o(current_l[1])
                + clabe(current_l[2])
                + banco_b(current_l[3])
                + nombre_b(current_l[4])
                + monto(current_l[5])
            )

            if with_cf == True:
                second_string = (
                    c_fiscal(current_l[6], with_cf)
                    + rfc_b(current_l[7], with_cf)
                    + iva(current_l[8], with_cf)
                    + ref_n(current_l[9])
                    + c_pago(current_l[10], with_cf)
                )
            else:
                second_string = (
                    ref_n(current_l[6])
                    + c_pago(current_l[7], with_cf)
                )

            final_str = (first_string + second_string)
            final_len = 150
            if with_cf == True:
                final_len = 177

            if len(final_str) == final_len:
                self.flist.append(final_str)
            else:
                break

    def tocsv(self):
        """modify self.flist to ideal list for output in csv format

        Returns:
            list: contains strings that follow format
        """
        # functions that modify each object to conform to format
        def cuenta_o(i_object):
            i_string = str(i_object)
            diff = 10 - len(i_string)
            return ('0' * diff) + i_string

        def rfc_o(i_object):
            return str(i_object)

        def clabe(i_object):
            i_string = str(i_object)
            diff = 20 - len(i_string)
            return ('0' * diff) + i_string

        def banco_b(i_object):
            i_string = str(i_object)
            diff = 3 - len(i_string)
            return ('0' * diff) + i_string

        def nombre_b(i_object):
            return str(i_object)

        def monto(i_object):
            return "{:.2f}".format(i_object)

        def c_fiscal(i_object, con_cf):
            i_string = str(i_object)
            if con_cf == True:
                if i_string == 'S':
                    return i_string
                else:
                    return 'S'
            else:
                pass

        def rfc_b(i_object, con_cf):
            if con_cf == True:
                return str(i_object)
            else:
                pass

        def iva(i_object, con_cf):
            if con_cf == True:
                return "{:.2f}".format(i_object)
            else:
                pass

        def ref_n(i_object):
            return str(i_object)

        def c_pago(i_object, con_cf):
            return str(i_object)

        for r in range(self.nrow):
            current_l = self.ilist[r]
            with_cf = bool(self.itype)

            first_string = (
                cuenta_o(current_l[0])
                + "," + rfc_o(current_l[1])
                + "," + clabe(current_l[2])
                + "," + banco_b(current_l[3])
                + "," + nombre_b(current_l[4])
                + "," + monto(current_l[5])
            )

            if with_cf == True:
                second_string = (
                    "," + c_fiscal(current_l[6], with_cf)
                    + "," + rfc_b(current_l[7], with_cf)
                    + "," + iva(current_l[8], with_cf)
                    + "," + ref_n(current_l[9])
                    + "," + c_pago(current_l[10], with_cf)
                )
            else:
                second_string = (
                    "," + ref_n(current_l[6])
                    + "," + c_pago(current_l[7], with_cf)
                )

            final_str = (first_string + second_string)
            final_len = 157
            if with_cf == True:
                final_len = 187

            if len(final_str) <= final_len:
                self.flist.append(final_str)
            else:
                break


class Main:
    """intialize program
    """

    def run(self, input_name, input_dataframe, format_type):
        """initialize general program

        Args:
            input_name (str): name of program
            input_dataframe (array): inputed file as array
            format_type (boolean): whether format of choice to output is txt/ csv
        """
        i_data = Check(input_dataframe)
        (i_type, p_format) = i_data.check_type()

        while True:
            if i_data.simple_check() == False:
                break
            if p_format == False:
                break

            output = Format(input_dataframe, i_type, i_data.nrow, i_data.ncol)
            if format_type == True:
                output.totxt()
                out_file = output.flist
                f_filename = input_name + '.txt'
            else:
                output.tocsv()
                out_file = output.flist
                f_filename = input_name + '.csv'

            with open(f_filename, 'w+') as ffile:
                for i in range(i_data.nrow):
                    string = out_file[i]
                    ffile.write(string + "\n")
                break

    def __init__(self, file_directory, compiling_type):
        """intialize whole program
        """
        self.i_fname = file_directory
        self.file_name = self.i_fname.replace(".xlsx", "")
        self.f_type = compiling_type

        try:
            self.df = read_excel(self.i_fname)
            self.run(self.file_name, self.df, self.f_type)
            self.message = "Program ran successfully"
        except:
            self.message = "Failed to operate"


class GUI:
    def __init__(self, master):
        self.master = master

        # create fixed window in the middle of monitor (hopefully)
        master.geometry("630x440+400+200")
        master.resizable(False, False)
        master.title("CAE para SPEI")
        self.canvas = Canvas(
            master,
            bg='#EAEDED'
        )
        self.canvas.place(
            relx=0, rely=0,
            relwidth=1, relheight=1
        )

        # attributes that would be used later down the line
        self.input_directory = ""
        self.input_file_name = ""
        self.compile_type = bool()

    def lpanel(self):
        """Method that initiates left panel of application
        """
        lp_bgcolor = '#E3E3E3'

        # create frame on which all modules of the left panel sit
        self.panel_left = Frame(
            self.canvas,
            bg=lp_bgcolor,
            bd=8
        )
        self.panel_left.place(
            relwidth=0.55, relheight=1
        )

        def alltitle():
            """creates both spanish title and 
            english sub-titles widgets with label
            """
            Label(
                self.panel_left,
                text="Compilador de Archivos\nExcel para SPEI",
                font=("default, 30"),
                bg=lp_bgcolor, fg="#000000",
                justify=LEFT
            ).place(
                relwidth=1, relheight=0.185
            )

            Label(
                self.panel_left,
                text="Excel File (.xlsx) Compiler for SPEI",
                font=("default, 21"),
                bg=lp_bgcolor, fg="#000000",
                justify=LEFT
            ).place(
                rely=0.195,
                relwidth=1, relheight=0.085
            )

        def espdef():
            """creates spanish definition text widget
            """
            self.espdef_box = LabelFrame(
                self.panel_left,
                bg=lp_bgcolor, fg="#000000",
                text="Descripción-ES"
            )
            self.espdef_box.place(
                rely=0.31,
                relwidth=1, relheight=0.355
            )

            self.espdef = Text(
                self.espdef_box,
                font=("default, 16"),
                highlightthickness=0,
                bg=lp_bgcolor, fg="#000000"
            )
            self.espdef.insert(
                END, "Esta aplicación esta hecha para crear\n")
            self.espdef.insert(
                END, "archivo tipo texto(.txt) o CSV(.csv.), con\n")
            self.espdef.insert(
                END, "información de archivos importada en\n")
            self.espdef.insert(
                END, "excel (.xlsx), que esta en el formato\n")
            self.espdef.insert(
                END, "para uso en el Sistema de Pagos Electró-\n")
            self.espdef.insert(
                END, "nicos Interbancarios (SPEI) de México.")
            self.espdef.config(state=DISABLED)
            self.espdef.place(
                relx=0.03, rely=0.033,
                relwidth=0.95, relheight=0.9
            )

        def engdef():
            """creates english definition text widget
            """
            self.engdef_box = LabelFrame(
                self.panel_left,
                bg=lp_bgcolor, fg="#000000",
                text="Description-EN"
            )
            self.engdef_box.place(
                rely=0.685,
                relwidth=1, relheight=0.315
            )

            self.engdef = Text(
                self.engdef_box,
                font=("default, 16"),
                highlightthickness=0,
                bg=lp_bgcolor, fg="#000000"
            )
            self.engdef.insert(
                END, "This application is made to create a file in\n")
            self.engdef.insert(
                END, "text (.txt) or CSV (.csv) format, with data\n")
            self.engdef.insert(
                END, "from an inputed excel (.xlsx) file, in set\n")
            self.engdef.insert(
                END, "format for use in the electronic interbank\n")
            self.engdef.insert(
                END, "payment system (SPEI) of Mexico.")
            self.engdef.config(state=DISABLED)
            self.engdef.place(
                relx=0.03, rely=0.035,
                relwidth=0.95, relheight=0.9
            )

        # initiates all "sub-widgets"
        alltitle()
        espdef()
        engdef()

    def rpanel(self):
        """
        Method that initiates right panel of application
        """
        rp_bgcolor = '#F0F0F0'

        # create frame on which all modules of the right panel sit
        self.panel_right = Frame(
            self.canvas,
            bg=rp_bgcolor,
            bd=8
        )
        self.panel_right.place(
            relx=0.55,
            relwidth=0.45, relheight=1
        )

        def instruction():
            """creates both spanish and english instruction sets
            """
            def esp():
                """spanish instruction set
                """
                self.tins_esp = Label(
                    self.panel_right,
                    text="Instrucciones-ES:",
                    font=("default, 19"),
                    bg=rp_bgcolor, fg="#000000",
                    anchor='w'
                )
                self.tins_esp.place(
                    rely=0.0235,
                    relwidth=1, relheight=0.07
                )

                self.esp_instruct = Text(
                    self.panel_right,
                    font=("default, 14"),
                    highlightthickness=0,
                    bg=rp_bgcolor, fg="#000000"
                )
                self.esp_instruct.insert(
                    END, "1 Haz click en 'Choose' y escoge archivo")
                self.esp_instruct.insert(
                    END, "\n2 Selecciona tipo de archivo de salida")
                self.esp_instruct.insert(
                    END, "\n3 Haz click en 'Compile'")
                self.esp_instruct.config(state=DISABLED)
                self.esp_instruct.place(
                    rely=0.0785,
                    relwidth=1, relheight=0.13
                )

            def eng():
                """english instruction set
                """
                self.tins_eng = Label(
                    self.panel_right,
                    text="Instructions-EN:",
                    font=("default, 19"),
                    bg=rp_bgcolor, fg="#000000",
                    anchor='w'
                )
                self.tins_eng.place(
                    rely=0.2135,
                    relwidth=1, relheight=0.07
                )

                self.eng_instruct = Text(
                    self.panel_right,
                    font=("default, 14"),
                    highlightthickness=0,
                    bg=rp_bgcolor, fg="#000000"
                )
                self.eng_instruct.insert(
                    END, "1 Click on 'Choose' and choose file")
                self.eng_instruct.insert(
                    END, "\n2 Choose output file type")
                self.eng_instruct.insert(
                    END, "\n3 Click 'Compile'")
                self.eng_instruct.config(state=DISABLED)
                self.eng_instruct.place(
                    rely=0.2685,
                    relwidth=1, relheight=0.13
                )

            esp()
            eng()

        def action_module():
            """create action module, initiates processes
            and is the main module which USERS interact with
            """
            def input_diplay_update(message):
                self.input_display.config(state=NORMAL)
                self.input_display.delete(1.0, END)
                self.input_display.insert(INSERT, message)
                self.input_display.config(state=DISABLED)

            def compile_diplay_update(message, clear_box):
                self.compile_display.config(state=NORMAL)
                if clear_box == True:
                    self.compile_display.delete(1.0, END)
                self.compile_display.insert(END, message + "\n")
                self.compile_display.config(state=DISABLED)

            def input_button_function():
                """input button function
                """
                filedirectory = askopenfile(filetypes=[('*', '*.xlsx')])
                if filedirectory is not None:
                    self.input_directory = str(filedirectory.name)
                    self.input_file_name = os.path.basename(
                        self.input_directory)

                    input_diplay_update(self.input_file_name)
                    try:
                        self.dataframe = read_excel(self.input_directory)
                        checked_df = Check(self.dataframe)
                        (self.input_typeCF, pass_format_bool) = checked_df.check_type()
                        simple_check_boolean = checked_df.simple_check()
                    except:
                        simple_check_boolean = None
                        pass_format_bool = None

                    if simple_check_boolean == False:
                        compile_diplay_update(
                            "Input set failed simple format check.\nChoose another file", False)
                    elif pass_format_bool == False:
                        compile_diplay_update(
                            "Input set has some values\nthat are too large.\nChoose another file", False)

                    if simple_check_boolean == True and pass_format_bool == True:
                        self.choice_txt.config(state=NORMAL)
                        self.choice_csv.config(state=NORMAL)
                    else:
                        self.choice_txt.config(state=DISABLED)
                        self.choice_csv.config(state=DISABLED)
                        self.compile_button.config(state=DISABLED)

                else:
                    input_diplay_update(self.input_file_name)

            comp_type = StringVar()

            def choice_rbutton_function():
                user_choice = comp_type.get()
                if user_choice == "txt":
                    self.compile_type = True
                elif user_choice == "csv":
                    self.compile_type = False

                self.compile_button.config(state=NORMAL)

            def compile_button_function():
                directory = self.input_directory
                compile_type = self.compile_type

                if compile_type == True or compile_type == False:
                    process = Main(directory, compile_type)
                    compile_diplay_update(process.message, True)
                else:
                    compile_diplay_update("Error happened", True)

            # create frame on which all action modules sit
            self.am_panel = LabelFrame(
                self.panel_right,
                bg=rp_bgcolor, fg="#000000",
                text="Action Module"
            )
            self.am_panel.place(
                relx=0, rely=0.45,
                relwidth=1, relheight=0.55
            )

            def am_buttons():
                """method that contains the code that would create both 'input' 
                and 'compile' buttons and a mini display showing chosen file
                """
                # input button
                self.input_button = Button(
                    self.am_panel,
                    text="Choose",
                    font="default, 19",
                    bg="#FFFFFF", fg="#000000",
                    activebackground="#FFFFFF",
                    activeforeground="#696969",
                    highlightbackground=rp_bgcolor,
                    command=input_button_function
                )
                self.input_button.place(
                    relx=0.05, rely=0.02,
                    relwidth=0.31, relheight=0.195
                )

                # display where file chosen is shown
                self.input_display = Text(
                    self.am_panel,
                    font=("default, 15"),
                    height=1,
                    highlightthickness=0,
                    bg="#FFFFFF", fg="#696969",
                    padx=10, pady=11,
                    state=DISABLED
                )
                self.input_display.place(
                    relx=0.395, rely=0.025,
                    relwidth=0.548, relheight=0.182
                )

                # compile button
                self.compile_button = Button(
                    self.am_panel,
                    text="Compile",
                    font="default, 19",
                    bg="#FFFFFF", fg="#000000",
                    activebackground="#FFFFFF",
                    activeforeground="#696969",
                    highlightbackground=rp_bgcolor,
                    command=compile_button_function,
                    state=DISABLED
                )
                self.compile_button.place(
                    relx=0.05, rely=0.41,
                    relwidth=0.9, relheight=0.195
                )

                # display messages to user
                self.compile_display = Text(
                    self.am_panel,
                    font=("default, 13"),
                    height=1,
                    highlightthickness=0,
                    bg="#FFFFFF", fg="#696969",
                    padx=5, pady=2,
                    state=DISABLED
                )
                self.compile_display.place(
                    relx=0.05515, rely=0.662,
                    relwidth=0.8897, relheight=0.28
                )

            def choice_radiobutton():
                """method that contains radio buttons
                """
                # frame in which both radio buttons sit on
                self.choice_frame = Frame(
                    self.am_panel,
                    bg=rp_bgcolor
                )
                self.choice_frame.place(
                    relx=0.15, rely=0.215,
                    relwidth=0.7, relheight=0.195
                )

                # txt
                self.choice_txt = Radiobutton(
                    self.choice_frame,
                    text=".txt",
                    font=("default, 18"),
                    bg=rp_bgcolor,
                    variable=comp_type, value="txt",
                    command=choice_rbutton_function,
                    state=DISABLED
                )
                self.choice_txt.place(
                    relwidth=0.4, relheight=1
                )

                # csv
                self.choice_csv = Radiobutton(
                    self.choice_frame,
                    text=".csv",
                    font=("default, 18"),
                    bg=rp_bgcolor,
                    variable=comp_type, value="csv",
                    command=choice_rbutton_function,
                    state=DISABLED
                )
                self.choice_csv.place(
                    relx=0.6,
                    relwidth=0.4, relheight=1
                )

            am_buttons()
            choice_radiobutton()

        instruction()
        action_module()


def main():
    root = Tk()
    gui = GUI(root)
    gui.lpanel()
    gui.rpanel()
    root.mainloop()


if __name__ == '__main__':
    main()
