import tkinter
import tkinter.messagebox
import customtkinter
from tkinter import filedialog

customtkinter.set_appearance_mode("System")  # Modes: "System" (standard), "Dark", "Light"
customtkinter.set_default_color_theme("blue")  # Themes: "blue" (standard), "green", "dark-blue"


class App(customtkinter.CTk):
    def __init__(self):
        super().__init__()

        # configure window
        self.title("DAILY REPORT EXCEL AUTOMATION")
        self.window_width = 1100
        self.window_height = 540
        self.geometry(self.middle_position(self.window_width, self.window_height))
        self.minsize(self.window_width, self.window_height)
      



        # configure grid layout (4x4)
        self.grid_columnconfigure(1, weight=1)
        self.grid_columnconfigure((2, 3), weight=0)
        self.grid_rowconfigure((0, 1), weight=3)  # 0.75 * 4 = 3
        self.grid_rowconfigure(2, weight=1)  # 0.25 * 4 = 1

        # create sidebar frame with widgets
        self.sidebar_frame = customtkinter.CTkFrame(self, width=140, corner_radius=0)
        self.sidebar_frame.grid(row=0, column=0, rowspan=4, sticky="nsew")
        self.sidebar_frame.grid_rowconfigure(4, weight=1)
        self.logo_label = customtkinter.CTkLabel(self.sidebar_frame, text="Automation Report", font=customtkinter.CTkFont(size=20, weight="bold"))
        self.logo_label.grid(row=0, column=0, padx=20, pady=(20, 10))
        #Sử dụng lambda sẽ trì hoãn việc thực thi sidebar_button_event cho đến khi nút thực sự được nhấp vào. Thay vì gọi trực tiếp khi tạo nút
        self.sidebar_button_1 = customtkinter.CTkButton(self.sidebar_frame, command=lambda: self.sidebar_button_event(True))
        self.sidebar_button_1.grid(row=1, column=0, padx=20, pady=10)
        self.sidebar_button_2 = customtkinter.CTkButton(self.sidebar_frame, command=lambda: self.sidebar_button_event(False))
        self.sidebar_button_2.grid(row=2, column=0, padx=20, pady=10)
        # self.sidebar_button_3 = customtkinter.CTkButton(self.sidebar_frame, command=self.sidebar_button_event)
        # self.sidebar_button_3.grid(row=3, column=0, padx=20, pady=10)
        self.appearance_mode_label = customtkinter.CTkLabel(self.sidebar_frame, text="Appearance Mode:", anchor="w")
        self.appearance_mode_label.grid(row=5, column=0, padx=20, pady=(10, 0))
        self.appearance_mode_optionemenu = customtkinter.CTkOptionMenu(self.sidebar_frame, values=["Light", "Dark", "System"],
                                                                       command=self.change_appearance_mode_event)
        self.appearance_mode_optionemenu.grid(row=6, column=0, padx=20, pady=(10, 10))
        self.scaling_label = customtkinter.CTkLabel(self.sidebar_frame, text="UI Scaling:", anchor="w")
        self.scaling_label.grid(row=7, column=0, padx=20, pady=(10, 0))
        self.scaling_optionemenu = customtkinter.CTkOptionMenu(self.sidebar_frame, values=["80%", "90%", "100%", "110%", "120%"],
                                                               command=self.change_scaling_event)
        self.scaling_optionemenu.grid(row=8, column=0, padx=20, pady=(10, 20))

        # # create main entry and button
        # self.entry = customtkinter.CTkEntry(self, placeholder_text="CTkEntry")
        # self.entry.grid(row=3, column=1, columnspan=2, padx=(20, 0), pady=(20, 20), sticky="nsew")

        # self.main_button_1 = customtkinter.CTkButton(master=self, fg_color="transparent", border_width=2, text_color=("gray10", "#DCE4EE"))
        # self.main_button_1.grid(row=3, column=3, padx=(20, 20), pady=(20, 20), sticky="nsew")

        # create textbox
        self.textbox = customtkinter.CTkTextbox(self, width=250)
        self.textbox.grid(row=0, column=1, padx=(20, 0), pady=(20, 0), sticky="nsew")


       

        
        # create scrollable frame
        self.scrollable_frame = customtkinter.CTkScrollableFrame(self, label_text="List Sheets")
        self.scrollable_frame.grid(row=0, column=2, padx=(20, 0), pady=(20, 0), sticky="nsew")
        self.scrollable_frame.grid_columnconfigure(0, weight=1)
        self.scrollable_frame_switches = []
        for i in range(100):
            switch = customtkinter.CTkSwitch(master=self.scrollable_frame, text=f"CTkSwitch {i}")
            switch.grid(row=i, column=0, padx=10, pady=(0, 20))
            self.scrollable_frame_switches.append(switch)

        # create checkbox and switch frame
        self.checkbox_slider_frame = customtkinter.CTkFrame(self)
        self.checkbox_slider_frame.grid(row=0, column=3, padx=(20, 20), pady=(20, 0), sticky="nsew")
        self.checkbox_1 = customtkinter.CTkCheckBox(master=self.checkbox_slider_frame)
        self.checkbox_1.grid(row=1, column=0, pady=(20, 0), padx=20, sticky="n")
        self.checkbox_2 = customtkinter.CTkCheckBox(master=self.checkbox_slider_frame)
        self.checkbox_2.grid(row=2, column=0, pady=(20, 0), padx=20, sticky="n")
        self.checkbox_3 = customtkinter.CTkCheckBox(master=self.checkbox_slider_frame)
        self.checkbox_3.grid(row=3, column=0, pady=20, padx=20, sticky="n")


        #CREATE A BIGGER LOAD DATA BUTTON IN ROW 1 AND COLUMN 1+2+3
        self.load_data_button = customtkinter.CTkButton(self, text="LOAD DATA", command=lambda: self.load_data_event)
        self.load_data_button.grid(row=1, column=1, columnspan=1, padx=(20, 0), pady=(30, 30), sticky="nsew")

        #CREATE A BIGGER PROCESS BUTTON IN ROW 1 AND COLUMN 1+2+3
        self.process_button = customtkinter.CTkButton(self, text="PROCESS", command=self.process_button_event)
        self.process_button.grid(row=1, column=2, columnspan=2, padx=(20, 20), pady=(30, 30), sticky="nsew")

        # set default values
        self.sidebar_button_1.configure(text="Data Path")
        self.sidebar_button_2.configure(text="Report Path")
        
        # self.scrollable_frame_switches[0].select()
        # self.scrollable_frame_switches[4].select()
        self.appearance_mode_optionemenu.set("Dark")
        self.scaling_optionemenu.set("100%")
        
        #default checkbox
        self.checkbox_1.configure(text="Debit")
        self.checkbox_2.configure(text="Cost")
        self.checkbox_3.configure(text="Daily")
        self.checkbox_1.select()
        self.checkbox_2.select()
        self.checkbox_3.select()



        self.textbox.insert("0.0", "Logging...\n\n")
        #disable user input
        self.textbox.configure(state="disabled")


    def insert_to_log(self, message):
        self.textbox.configure(state="normal")
        self.textbox.insert(text= f" {message}\n\n", index="end")
        self.textbox.configure(state="disabled")

    def change_appearance_mode_event(self, new_appearance_mode: str):
        customtkinter.set_appearance_mode(new_appearance_mode)

    def change_scaling_event(self, new_scaling: str):
        new_scaling_float = int(new_scaling.replace("%", "")) / 100
        customtkinter.set_widget_scaling(new_scaling_float)

    def sidebar_button_event(self,data_button: False):
        

        if data_button:
            data_path = filedialog.askopenfilename(initialdir = "/",
                                                title = "Select a Data File",
                                                filetypes=(("Excel Files", "*.xlsx"),)) 
            if data_path == "":
                return
            # self.sidebar_button_1.configure(text=data_path)  # Update button text
            self.data_path = data_path
            #append to last line of textbox for logging
            self.insert_to_log(f" Sucess Loading Data Path: {data_path}")


        else:
            report_path = filedialog.askopenfilename(initialdir = "/",
                                                    title = "Select a Report File",
                                                    filetypes=(("Excel Files", "*.xlsx"),)) 
            if report_path == "":
                return
            # self.sidebar_button_2.configure(text=report_path)  # Update button text
            self.report_path = report_path
            #append to textbox for logging
            self.insert_to_log(f" Sucess Loading Report Path: {report_path}")
        
    
    def middle_position(self,window_width, window_height):
        screen_width = self.winfo_screenwidth()
        screen_height = self.winfo_screenheight()
        
        x_coordinate = (screen_width // 2) - (window_width // 2)
        y_coordinate = (screen_height // 2) - (window_height // 2)
        return f'{window_width}x{window_height}+{x_coordinate}+{y_coordinate}'

    def load_data_event(self):
        pass
    def process_button_event(self):
        pass


if __name__ == "__main__":
    app = App()
    app.mainloop()