import selenium.common.exceptions as exceptions  # To use SessionNotCreatedException from selenium
from selenium import webdriver  # To use the ChromeDriver
from selenium.webdriver.chrome.options import Options  # To set selenium options
from tkinter import *  # For all GUI building
from tkinter import messagebox  # Pop-up message
from tkinter import ttk  # To enable styling
from tkinter import filedialog  # To enable user to save export file in desired location
import ttkthemes as themes  # Styling package
import webbrowser  # To enable link opening
import mysql.connector as mysql  # To enable database interactions
import csv  # To write the export file
import time  # To get current date and time to name CSV file

global keyword_combobox, website_combobox


class MainWindow:
    # Constructor
    def __init__(self):
        # Initialisations
        global keyword_combobox, website_combobox
        self.search_topics = []
        self.search_sites = []
        cur = None
        db = None

        # Trying to connect to database, throw error if client did not start the MYSQL database
        try:
            db = mysql.connect(host="localhost", user="Client", passwd="1234", database="article_finder")
            cur = db.cursor()
        except mysql.Error as e:
            messagebox.showinfo("Database Connection", "FAILED! Could not connect to MYSQL server. " + str(e))
            exit()

        # Getting current keywords, putting them into search_topics list
        cur.execute("SELECT topic from keywords")
        search_topics = [x[0] for x in cur.fetchall()]

        # Getting current websites, putting them into search_links list
        cur.execute("SELECT link from websites")
        search_links = [x[0] for x in cur.fetchall()]
        db.close()  # Closing db connection as we are done for now

        # Initialise and set up window size and theme package
        root = Tk()
        root.title("Article Finder")
        self.screen_width = root.winfo_screenwidth()
        self.screen_height = root.winfo_screenheight()
        root.geometry(f'{self.screen_width}x{self.screen_height}')  # Take whole screen
        root_style = themes.ThemedStyle(root)
        root_style.set_theme("arc")  # equilux

        # Create top frame and grid into root
        frame_top = ttk.Frame(root)
        frame_top.grid(row=0, column=0, sticky="nsew")

        # Create middle frame and grid into root
        frame_middle = ttk.Frame(root)
        frame_middle.grid(row=1, column=0, sticky="nsew")

        # Create bottom frame and grid into root
        frame_bottom = ttk.Frame(root)
        frame_bottom.grid(row=2, column=0, sticky="nsew")

        # Row and column configuration to allow stretch
        Grid.rowconfigure(root, 0, weight=2)  # top frame 1 gets 2/11 of the screen
        Grid.rowconfigure(root, 1, weight=8)  # middle frame 2 gets 8/11 of the screen
        Grid.rowconfigure(root, 2, weight=1)  # bottom frame gets 1/11 of the screen
        Grid.columnconfigure(root, 0, weight=1)

        # Create keywords label and combobox, grid into top frame
        KLabel = ttk.Label(frame_top, text="Keyword:")
        KLabel.grid(row=0, column=0)
        keyword_combobox = ttk.Combobox(frame_top, state="readonly")
        keyword_combobox["values"] = search_topics
        keyword_combobox.set(search_topics[0])
        keyword_combobox.grid(row=0, column=1)

        # Create websites label and combobox, grid into top frame
        WLabel = ttk.Label(frame_top, text="Website:")
        WLabel.grid(row=0, column=2)
        website_combobox = ttk.Combobox(frame_top, state="readonly")
        website_combobox["values"] = search_links
        website_combobox.set(search_links[0])
        website_combobox.grid(row=0, column=3)

        # Keywords button placed under the corresponding label and combobox, column span is 2
        keywords_button = ttk.Button(frame_top, text="Edit Keywords", command=lambda: EditWindow("Keywords"))
        keywords_button.grid(row=1, column=0, columnspan=2, padx=10, pady=3, sticky="nsew")

        # Websites button placed under the corresponding label and combobox, column span is 2
        websites_button = ttk.Button(frame_top, text="Edit Websites", command=lambda: EditWindow("Websites"))
        websites_button.grid(row=1, column=2, columnspan=2, padx=10, pady=3, sticky="nsew")

        # Search button calls Display Articles to search and fill up results table, placed under other buttons
        search_button = ttk.Button(frame_top, text="Search for Articles",
                                   command=lambda: self.display_articles(keyword_combobox.get(),
                                                                         website_combobox.get()))
        search_button.grid(row=2, column=0, columnspan=4, padx=10, pady=3, sticky="nsew")

        # Building results table (called TreeView)
        self.results_table = ttk.Treeview(frame_middle)
        ttk.Style().configure('Treeview', rowheight=35)
        self.results_table["columns"] = ("#0", "#1", "#2")
        self.results_table.column("#0", width=int(self.screen_width * 500 / 1440),
                                  minwidth=int(self.screen_width * 500 / 1440))
        self.results_table.column("#1", width=int(self.screen_width * 910 / 1440),
                                  minwidth=int(self.screen_width * 910 / 1440))
        self.results_table.column("#2", width=1, minwidth=1)

        # Setting Headings
        self.results_table.heading("#0", text="Title")
        self.results_table.heading("#1", text="Blurb")
        self.results_table.heading("#2", text="")

        self.results_table["displaycolumns"] = ("#0", "#1")  # Only Titles and Blurbs should be displayed
        self.results_table.bind("<Double-1>", self.open_articles)  # Binding double click
        self.results_table.grid(row=0, column=0, sticky="nsew", padx=10, pady=3)

        export_btn = ttk.Button(frame_bottom, text="Export Articles", command=self.export_articles)
        export_btn.grid(row=0, column=0, sticky="nsew", padx=10, pady=3)

        # Row and column configs for each frame to allow stretching inside each frame
        Grid.rowconfigure(frame_top, 0, weight=1)
        Grid.rowconfigure(frame_top, 1, weight=1)
        Grid.rowconfigure(frame_top, 2, weight=1)

        Grid.columnconfigure(frame_top, 0, weight=1)
        Grid.columnconfigure(frame_top, 1, weight=1)
        Grid.columnconfigure(frame_top, 2, weight=1)
        Grid.columnconfigure(frame_top, 3, weight=1)

        Grid.rowconfigure(frame_middle, 0, weight=1)
        Grid.columnconfigure(frame_middle, 0, weight=1)

        Grid.rowconfigure(frame_bottom, 0, weight=1)
        Grid.columnconfigure(frame_bottom, 0, weight=1)

        root.mainloop()

    # Puts results into table
    def display_articles(self, keyword, website):
        found_articles = Scraper.find_articles(keyword, website)  # 2D list of each row to use in the table
        self.results_table.delete(*self.results_table.get_children())  # Clear current tree
        # Inserting into tree
        for i in range(len(found_articles)):
            self.results_table.insert("", i, text=found_articles[i][0],
                                      values=(found_articles[i][1], found_articles[i][2]))

    # Opens clicked articles
    def open_articles(self, event):
        Item = self.results_table.selection()[0]  # First item in the user's selection. Only 1 link is opened
        url = self.results_table.item(Item, "values")[1]  # The url taken from the table
        webbrowser.open_new(url)  # Open URL

    # Updates chosen combobox
    @staticmethod
    def update_combobox(category):
        items = TreeAndDatabaseFunctions.search_database(category.lower())  # Pull table of keywords and websites
        newItems = []
        for item in items:
            newItems.append(item[1])  # Taking only the relevant values of the pulled table (Excluding "ID")

        #  Update the necessary combo box
        if category == "Keywords":
            keyword_combobox["values"] = newItems
            keyword_combobox.set(newItems[0])
        elif category == "Websites":
            website_combobox["values"] = newItems
            website_combobox.set(newItems[0])

    # Exports search results into a CSV file
    def export_articles(self):
        results = list(self.results_table.get_children())  # Everything in the tree
        to_export = []
        for result in results:
            temp = [self.results_table.item(result, "text")]  # Getting title in each tree row object
            for item in list(self.results_table.item(result, "values")):  # Getting blurbs and links in each tree row
                temp.append(item)  # temp has one row, complete at this stage
            to_export.append(temp)  # Appending temp to the bigger row
        keyword = keyword_combobox.get()  # Getting chosen keyword
        website = website_combobox.get().split(".")[0]  # Getting just the domain name of chosen website (ex. "wired")

        if len(to_export) == 0:  # Checking if results array is empty (if user clicked export without searching)
            messagebox.showinfo("Error", "Cannot export an empty table!")  # Popped error message
        else:
            date = "-".join([str(time.localtime()[x]) for x in range(0,3)])
            times = ".".join([str(time.localtime()[x]) for x in range(3,6)])
            file_name = filedialog.asksaveasfilename(title='Export Results', confirmoverwrite=True,
                                                     initialfile=f'Exported_{keyword}_{website} {date} at {times}',
                                                     defaultextension=".csv")  # Asking user for file name and save path
            file = open(f"{file_name}", "w")  # Opening file that user just made
            csvwriter = csv.writer(file)  # Creating writer, passing file. Since it is a CSV, commas are the delimiters
            csvwriter.writerow(["Title", "Blurb", "Link"])  # Writing headings as the first row
            for row in to_export:
                csvwriter.writerow(row)  # Writing each result row iteratively
            file.close()  # Closing file as we are done


class EditWindow:
    # Constructor
    def __init__(self, category):
        # The creating edit window
        self.edit_window = Tk()
        self.edit_window.geometry("410x600")
        self.edit_window.title(f"Edit {category}")
        self.edit_window.resizable = False

        # Set theme
        EditWin_Style = themes.ThemedStyle(self.edit_window)
        EditWin_Style.set_theme("arc")  # equilux

        # Initialising frames, tree and label
        f1 = ttk.Frame(self.edit_window)
        f1.pack(fill=BOTH, expand=True)
        f2 = ttk.Frame(self.edit_window)
        f2.pack(fill=BOTH, expand=True)
        CurrentLabel = ttk.Label(f1, text=f"Current {category}", font=("Roboto", 25))
        CurrentLabel.pack(anchor='w', padx=5, pady=5)
        self.current_table = ttk.Treeview(f1)

        # Setting up tree
        self.current_table['columns'] = "#1"
        self.current_table.column("#0", width=60)
        self.current_table.column("#1", width=335)
        self.current_table.heading("#0", text="ID")
        self.current_table.heading("#1", text=f"{category}")
        self.current_table.bind("<BackSpace>", lambda event: self.delete_keyword_or_website(category))
        self.current_table.pack(expand=TRUE, fill=BOTH, padx=5, pady=5)

        # Pulls all rows of either category from database and stores in array
        KeywordsOrSitesTable = TreeAndDatabaseFunctions.search_database(f"{category.lower()}")

        # Insert into Current_Tree table
        for i in range(len(KeywordsOrSitesTable)):
            self.current_table.insert("", i, text=f"{KeywordsOrSitesTable[i][0]}", values=(KeywordsOrSitesTable[i][1:]))

        # Entry box
        self.entry_box = ttk.Entry(f2)
        self.entry_box.bind("<Return>",
                            lambda event: self.add_keyword_or_website(category))  # Bind return key to Add func
        self.entry_box.grid(row=0, column=0, sticky="nsew", padx=5, pady=5)

        # Buttons
        AddBtn = ttk.Button(f2, text=f"Add {category[:len(category) - 1]}",
                            command=lambda: self.add_keyword_or_website(category))
        AddBtn.grid(row=0, column=1, sticky="nsew", padx=5, pady=5)

        DelBtn = ttk.Button(f2, text=f"Remove Selected {category}",
                            command=lambda: self.delete_keyword_or_website(category))
        DelBtn.grid(row=1, column=0, sticky="nsew", padx=5, pady=5, columnspan=2)

        QuitBtn = ttk.Button(f2, text="Back to Search", command=lambda: self.edit_window.destroy())
        QuitBtn.grid(row=2, column=0, sticky="nsew", padx=5, pady=5, columnspan=2)

        # Row and column configs to allow stretching
        Grid.rowconfigure(f2, 0, weight=1)
        Grid.rowconfigure(f2, 1, weight=1)
        Grid.rowconfigure(f2, 2, weight=1)
        Grid.columnconfigure(f2, 0, weight=1)
        Grid.columnconfigure(f2, 1, weight=1)

    # Allows user to add keywords or websites
    def add_keyword_or_website(self, choice):
        col_head = ""
        if choice == "Keywords":
            col_head = "topic"  # Specific column title in database
        elif choice == "Websites":
            col_head = "link"  # Specific column title in database

        db = mysql.connect(host="localhost", user="Client", passwd="1234", database="article_finder")  # Connect to db
        cur = db.cursor()
        query = f"INSERT INTO {choice.lower()} ({col_head}) VALUES ('{self.entry_box.get()}')"
        cur.execute(query)  # inserting into the table
        db.commit()  # save changes
        db.close()  # close connection for now
        self.entry_box.delete(0, END)  # clear entry box
        TreeAndDatabaseFunctions.update_tree(self.current_table, choice)  # update Current_Tree with new values
        MainWindow.update_combobox(choice)

    # Allows user to delete keywords or websites
    def delete_keyword_or_website(self, choice):
        Selected = list(self.current_table.selection())  # Get the rows that the user selected
        SelectedIDs = [self.current_table.item(Item, "text") for Item in Selected]  # Put the ids of each row in array
        db = mysql.connect(host="localhost", user="Client", passwd="1234", database="article_finder")  # Connect db
        cur = db.cursor()
        # Deleting each row by id from SelectedIds array
        for ID in SelectedIDs:
            query = f"DELETE FROM {choice.lower()} WHERE id = {ID}"
            cur.execute(query)
        db.commit()  # Save changes
        db.close()  # Close database connection for now
        TreeAndDatabaseFunctions.update_tree(self.current_table, choice)  # Update table of keywords/websites
        MainWindow.update_combobox(choice)  # Update keyword/website combobox based on what was changed


class TreeAndDatabaseFunctions:
    # Returns all rows in the given (keyword or website) table
    @staticmethod
    def search_database(category):
        db = mysql.connect(host="localhost", user="Client", passwd="1234", database="article_finder")
        cur = db.cursor()
        cur.execute(f"SELECT * FROM {category}")  # Select all rows from specified table
        items = cur.fetchall()
        db.close()

        final = []
        # Place each row as a list in the "final" list
        for item in items:
            final.append(list(item))  # Convert from tuple to list and append to final list
        return final

    # Updates a passed keyword/website table with latest data
    @staticmethod
    def update_tree(tree, category):
        tree.delete(*tree.get_children())  # Clear tree
        items = TreeAndDatabaseFunctions.search_database(category)  # Get latest data
        for i in range(len(items)):
            tree.insert("", i, text=items[i][0], values=items[i][1:])  # Insert into table


class Scraper:
    # Finds articles
    @staticmethod
    def find_articles(keyword, website):
        options = Options()  # Selenium options
        options.add_argument('--headless')  # No Chrome popup
        driver = None  # Initialize driver
        try:
            driver = webdriver.Chrome(executable_path="/usr/local/bin/chromedriver",
                                      options=options)  # Trying to create driver

        except exceptions.SessionNotCreatedException as e:  # This is an exception from Selenium
            messagebox.showinfo("Please Update ChromeDriver",  # If wrong version, throw error and give url for download
                                f"{str(e)[30:132]}\n\n"  # relevant section of Selenium exception
                                f"Please RE-DOWNLOAD the correct ChromeDriver version from:\n"
                                f"https://chromedriver.chromium.org/downloads\n\n"  # Chrome driver website
                                f"REPLACE 'ChromeDriver' executable in the following path with the new one:\n"
                                f"\n/usr/local/bin\n"
                                f"\n⌘+⇧+G on Finder (use above path)")  # Correct path to place ChromeDriver for program
            exit()  # Kill program

        except exceptions.WebDriverException:  # This is another exception from Selenium
            messagebox.showinfo("ChromeDriver Not Found",  # If ChromeDriver is in the wrong path, throw error
                                "ChromeDriver needs to be moved to:\n"
                                "\n/usr/local/bin\n"  # Required (Target) Path
                                "\n⌘+⇧+G on Finder (use above path)")  # Instruction for user
            exit()  # Kill program

        target = f"https://www.google.com/search?hl=en&q={keyword}+site:{website}&tbm=nws&tbs=qdr:w"
        driver.get(target)
        driver.implicitly_wait(3)  # Ensure page load

        # Getting the titles of articles
        raw_titles = driver.find_elements_by_class_name("JheGif")  # gets objects
        titles = []  # to put title strings into array
        for title in raw_titles:
            titles.append(title.text)

        # Getting the blurbs of articles
        raw_blurbs = driver.find_elements_by_class_name("Y3v8qd")
        blurbs = []
        for blurb in raw_blurbs:
            blurbs.append(blurb.text)

        # Getting the links of articles, always in "a" tag within the "dbsr" class
        elements = driver.find_elements_by_class_name("dbsr")  # Finds "dbsr" classes that contain the URL tags
        links = []
        for a_tag in elements:  # iterating through each dbsr class
            child = a_tag.find_element_by_css_selector("*")  # finds the first of each dbsr tag's children (always URL)
            link = child.get_attribute("href")  # getting href attribute from the a tag child
            links.append(link)

        driver.quit()

        # Final processing into 2D Array for MainWindow class
        found_articles = []
        for x in range(len(titles)):
            found_articles.append([titles[x], blurbs[x], links[x]])

        return found_articles


MainWindow()
