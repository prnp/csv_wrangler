#!/usr/bin/python

###########################################################
# Author: Randy Braun (braun.randy@gmail.com)
# 
# CSV wrangler is a simple gui tool that can be used to easily 
#  trim down delimited files that contain a large number of
#  of fields. Simply choose the file to load, the fields you
#  would like to extract, choose the file you want to save
#  the fields in, and go! 
#
###########################################################

import tkFileDialog as tfd
import tkMessageBox as box
import Tkconstants, ttk

import csv

import Tkinter as tk
import sys

class csv_wrangler():

    def __init__(self):
        """constructor"""
        self.left_width = 40
        self.right_width = 40
        self.column_delimiter = '|'
        self.header_dict = {}
        self.header_list = []
        self.default_template = '88,109,55,112,57,70,114,1,4,8,15,87'        
        self.__init_gui()
        

    def run(self):
        """method to run the main loop"""
        self.root.mainloop()

    def __move_column(self,from_listbox, to_listbox,dec_count,inc_count):
        """
        function to items between listboxes
        """
        to_move = from_listbox.curselection()

        try:
            # insert into new list
            for index in to_move:
                # get the line's text
                seltext = from_listbox.get(index)
                # insert into the new list
                to_listbox.insert(tk.END, seltext)
            
            # go cleanup and remove
            for index in reversed(to_move):
                # remove the line from the old list
                from_listbox.delete(index)
                # decrement the counter
                dec_tmp = int(dec_count.get()) - 1
                dec_count.set("%d" % dec_tmp)
                # increment the counter
                inc_tmp = int(inc_count.get()) + 1
                inc_count.set("%d" % inc_tmp)
                
        except IndexError as detail:
            box.showinfo("Information", detail)


    def __askopenfile(self,name_var):
        """Sets name of file to open"""
        
        old_name = name_var.get()
        file_opt = {'defaultextension':'.csv',
                    'filetypes':[('CSV files','.csv')]
                    }

        file_name = tfd.askopenfilename(**file_opt)

        if len(file_name) > 0:
            name_var.set(file_name)
        else:
            name_var.set(old_name)

    def __asksavefile(self,name_var):
        """Sets name of file to save to"""
        
        old_name = name_var.get()
        file_opt = {'defaultextension':'.csv',
                    'filetypes':[('CSV files','.csv')]
                    }

        file_name = tfd.asksaveasfilename(**file_opt)

        if len(file_name) > 0:
            name_var.set(file_name)
        else:
            name_var.set(old_name)

    def __get_headers(self):
        """Function to retrieve the header columns from the csv file"""
        rv = 0
        self.column_delimiter = self.col_delim.get()
        file_name = self.load_entry.get()
        try:
            with open(file_name,'r') as fh:
                reader = csv.reader(fh, delimiter=self.column_delimiter, quoting=csv.QUOTE_NONE)
                # Get list that contains the column header names in appropriate order
                self.header_list = reader.next()
            rv = 1
        except IOError, detail:
            box.showinfo("Error", "Unable to open file: %s \n %s" % (file_name,detail))

        # Set the dictionary that maps column names to index
        #  We'll keep the 1st element indexed at 0 to make
        #  things simple and adjust appropriately when loading
        #  from the template field.
        for x in xrange(0,len(self.header_list)):
            self.header_dict[self.header_list[x]] = x
            
        return rv

#    def __set_headers_dict(self, import_file):
#        """Function to retrieve headers"""
#        self.column_delimiter = self.col_delim.get()
#        filename = import_file.get()
#        self.header_dict = scsv.get_header_dict(filename,self.column_delimiter)

#    def __set_headers_list(self, import_file):
#        """Function to retrieve headers"""
#        self.column_delimiter = self.col_delim.get()
#        filename = import_file.get()
#        self.header_list = scsv.get_header_list(filename,self.column_delimiter)

    def __load_header(self,listbox, import_file):
        """Method to populate the listbox from the import"""
        #self.__set_headers_dict(import_file)
        self.__get_headers()
        # Remove current column names from the listbox
        listbox.delete(0,tk.END)

        for key in self.header_list:
            listbox.insert(self.header_dict[key],key)

        # Set number of columns
        self.import_count.set(len(self.header_list))

    def __template_ctl(self):
        """control form elements when toggling templates"""
        
        flag = self.template_load_flag.get()
        if flag > 0: #Enabled
            #Disable listbox and load button
            self.listbox1.config(state=tk.DISABLED)
            self.load_button.config(state=tk.DISABLED)
            self.button1.config(state=tk.DISABLED)
            self.button2.config(state=tk.DISABLED)
            self.go_button.config(state=tk.NORMAL)
            self.template_entry.config(state=tk.NORMAL)
        else: #Disabled
            self.listbox1.config(state=tk.NORMAL)
            self.load_button.config(state=tk.NORMAL)
            self.button1.config(state=tk.NORMAL)
            self.button2.config(state=tk.NORMAL)
            self.go_button.config(state=tk.DISABLED)
            self.template_entry.config(state=tk.DISABLED)

    def __load_template(self):
        """load fields from template into listbox2"""
        # Get value from template entry field
        field_txt = self.template_txt.get()
        # 
        if self.__get_headers():
            go_on = 1
        else:
            go_on = 0

        if go_on:
            # Remove current column names from the listbox
            self.listbox2.delete(0,tk.END)
            # Reset counter
            self.export_count.set(0)
            
            if len(field_txt) < 1:
                box.showinfo("Error", "The template field is empty")
                return
            field_list = field_txt.split(',')
            if len(field_list) < 1:
                box.showinfo("Error", "The field list is empty")
                return
            exp_cnt = int(self.export_count.get())
            for idx in field_list:
                try:
                    i = int(idx) - 1
                    self.listbox2.insert(tk.END,self.header_list[i])
                    exp_cnt += 1
                except Exception, detail:
                    box.showinfo("Error","Field %s, doesn't exist. %s" % (idx,detail))
            self.export_count.set(exp_cnt)

    def __write_file(self):
        """Callback method to write out the requested file"""
        # get fields to write form listbox2
        field_list = self.listbox2.get(0,tk.END)
        import_file = self.load_entry.get()
        export_file = self.save_entry.get()

        count = 0
        try:
            with open(export_file,'wb') as fw:
                with open(import_file,'r') as fr:
                    reader = csv.reader(fr, delimiter=self.column_delimiter,quoting=csv.QUOTE_NONE)
                    writer = csv.writer(fw, dialect='excel')
                    #Output the header row to file
                    writer.writerow(field_list)
                    first_row = 1
                    for row in reader:
                        if first_row:
                            first_row = 0
                            continue
                        count += 1
                        row_to_write = []
                        for col in field_list:
                            row_to_write.append(row[self.header_dict[col]])

                        #print row_to_write
                        writer.writerow(row_to_write)

            box.showinfo("information","Wrote %s lines to %s" % (count,export_file))
        except Exception, detail:
            box.showinfo("Error", "Error saving file: %s \n %s" % (export_file,detail))

    def __openSettings(self):
        """TODO: Open a new window to manipulate program settings"""
        settings_win = tk.Toplevel()
        settings_win.title("Application Settings")
	

    def __onExit(self):
        """Things we do before exiting the application"""
        quit()

    def __init_gui(self):
        """Function to setup window architecture"""

        self.root = tk.Tk()
        self.root.title("CSV Wrangler")
        self.import_count = tk.StringVar()
        self.export_count = tk.StringVar()

        # Create menubar
        menubar = tk.Menu(self.root)
        self.root.config(menu=menubar)
        filemenu = tk.Menu(menubar)
        filemenu.add_command(label="Settings", command=self.__openSettings)
        filemenu.add_command(label="Exit", command=self.__onExit)
        menubar.add_cascade(label="File", menu=filemenu)

        # left listbox labels
        self.lb_label1_txt = tk.StringVar()
        lb_label1 = tk.Label(self.root, fg='black',textvariable=self.lb_label1_txt)
        lb_label1.grid(row=0, column=0, sticky=tk.W)
        self.lb_label1_txt.set("Imported Columns:")

        #self.col_num = tk.StringVar()
        col_num_label = tk.Label(self.root, fg='black', textvariable=self.import_count)
        col_num_label.grid(row=0, column=1, sticky=tk.W)
        self.import_count.set("%d" % 0)

        # Right listbox labels
        self.lb_label2_txt = tk.StringVar()
        lb_label2 = tk.Label(self.root, fg='black',textvariable=self.lb_label2_txt)
        lb_label2.grid(row=0, column=6, sticky=tk.W)
        self.lb_label2_txt.set("Columns to Export:")

        #self.exp_col_num = tk.StringVar()
        exp_col_num_label = tk.Label(self.root, fg='black', textvariable=self.export_count)
        exp_col_num_label.grid(row=0, column=7, sticky=tk.W)
        self.export_count.set("%d" % 0)

        # create the listbox (note that size is in characters)
        self.listbox1 = tk.Listbox(self.root, width=self.left_width, height=6,selectmode=tk.EXTENDED)
        self.listbox1.grid(row=1, column=0, columnspan=4, rowspan=4)
        self.listbox2 = tk.Listbox(self.root, width=self.right_width, height=6,selectmode=tk.EXTENDED)
        self.listbox2.grid(row=1, column=6, columnspan=4, rowspan=4)

        # create a vertical scrollbar for the left listbox 
        yscroll1 = tk.Scrollbar(command=self.listbox1.yview, orient=tk.VERTICAL)
        yscroll1.grid(row=1, column=4, rowspan=4, sticky=tk.N+tk.S)
        self.listbox1.configure(yscrollcommand=yscroll1.set)

        # create a vertical scrollbar for the right listbox 
        yscroll2 = tk.Scrollbar(command=self.listbox2.yview, orient=tk.VERTICAL)
        yscroll2.grid(row=1, column=10, rowspan=4, sticky=tk.N+tk.S)
        self.listbox2.configure(yscrollcommand=yscroll2.set)

        # button to move items from right to left
        self.button1 = tk.Button(self.root, text='<--',
                            command=lambda: self.__move_column(self.listbox2,self.listbox1,
                                                               self.export_count,self.import_count))
        self.button1.grid(row=2, column=5, padx=10)

        # button to move items from left to right
        self.button2 = tk.Button(self.root, text='-->',
                            command=lambda: self.__move_column(self.listbox1,self.listbox2,
                                                        self.import_count,self.export_count))
        self.button2.grid(row=3, column=5)

        ### left side
        # left spacer label
        label_spacer1 = tk.Label(self.root, fg='red',text="")
        label_spacer1.grid(row=5, column=0)

        # label on left
        label1_txt = "File to load:"
        label1 = tk.Label(self.root, fg='black',text=label1_txt)
        label1.grid(row=6, column=0)

        # file load button
        self.load_button = tk.Button(self.root, text="Load",
                                command=lambda: self.__load_header(self.listbox1,
                                                                   self.load_entry))
        self.load_button.grid(row=6, column=1, columnspan=3, padx=5, sticky=tk.E)

        # file entry box left
        self.load_entry = tk.StringVar()
        file_load = tk.Entry(self.root, width=self.left_width, bg='yellow', textvariable=self.load_entry)
        file_load.grid(row=7, column=0, columnspan=5, pady=5)
        self.load_entry.set("No file selected")

        # file browse load button
        file_open = tk.Button(self.root, text="Browse", command=lambda: self.__askopenfile(self.load_entry))
        file_open.grid(row=6, column=1, sticky=tk.W)        

        # column delimiter setting
        delim_label_txt = "Column Delimiter:"
        delim_label = tk.Label(self.root,fg="black",text=delim_label_txt)
        delim_label.grid(row=8,column=0,sticky=tk.W, pady=5)
        
        self.col_delim = tk.StringVar()
        delim_entry = tk.Entry(self.root, width=2, textvariable=self.col_delim)
        delim_entry.grid(row=8,column=0,sticky=tk.E)
        self.col_delim.set(self.column_delimiter)

        # load from template checkbox
        self.template_load_flag = tk.IntVar()
        template_cb = tk.Checkbutton(self.root, variable=self.template_load_flag,
                                     text="Load from template:", command=self.__template_ctl)
        template_cb.grid(row=9,column=0, sticky=tk.W)

        # template entry box
        self.template_txt = tk.StringVar()
        self.template_entry = tk.Entry(self.root, width=20, bg='white',
                                       textvariable=self.template_txt)
        self.template_entry.insert(0,self.default_template)
        self.template_entry.config(state=tk.DISABLED)
        self.template_entry.grid(row=9,column=1, columnspan=4)

        # template GO button
        self.go_button = tk.Button(self.root, text="Go", state=tk.DISABLED,
                                   command=self.__load_template)
        self.go_button.grid(row=9,column=5, padx=0)

        ### Right side
        # Right spacer label
        label_spacer2 = tk.Label(self.root, fg='red',text="")
        label_spacer2.grid(row=5, column=6)

        # label on right
        label2_txt = "Save file as:"
        label2 = tk.Label(self.root, fg='black',text=label2_txt)
        label2.grid(row=6, column=6)

        # file save button
        save_button = tk.Button(self.root, text="Save", command=self.__write_file)
        save_button.grid(row=6, column=8, sticky=tk.W)

        # file save box right
        self.save_entry = tk.StringVar()
        file_save = tk.Entry(self.root, width=self.right_width, bg='yellow', textvariable=self.save_entry)
        file_save.grid(row=7, column=6, columnspan=5)
        self.save_entry.set("No file selected")

        # file browse save button
        file_save = tk.Button(self.root, text="Browse", command=lambda: self.__asksavefile(self.save_entry))
        file_save.grid(row=6, column=7, sticky=tk.W)

def main():

    csvw = csv_wrangler()
    csvw.run()

    


if __name__ == '__main__':
    main()
