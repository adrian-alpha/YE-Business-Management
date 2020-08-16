"""
Codename Genesis - 5674/87394.050-9

GENESIS BM

©2017-2019 Cortex™ - a direct subsidiary of Vortex Enterprise International™ - all rights reserved.
"""

try:
    import tkinter as tk
    from tkinter import ttk, messagebox, filedialog, scrolledtext, simpledialog
except ImportError:
    import Tkinter as tk
    from Tkinter import ttk, messagebox, filedialog, scrolledtext, simpledialog

paypal_mode = 'paypal_live'
version = '5.67'
website = "http://vortexenterpriseco.wixsite.com/vortex"

#user access levels:
# 1 = teacher/supervisor
# 2 = CEO
# 3 = Chief Officers
# 4 = Key Employees
# 5 = Basic employees
# 4 and 5 are just status', they do not differ in restrictions

class Main:
    def __init__(self, parent):
        #parent is the root window - Tk()
        self._top = parent
        parent.title("Starting Up...")
        #following lines catch if user tries to close the window
        parent.protocol('WM_DELETE_WINDOW', self.file_exit)
        #the icon for the root window: .xbm for linux; .ico for windows
        try:
            parent.wm_iconbitmap("Images/young enterprise 1.ico")
        except:
            parent.wm_iconbitmap("@Images/young enterprise 1.xbm")
        parent.geometry("800x600")
        self.showDebug = tk.StringVar()

        #accelerators are the shortcuts
        self.accelerators(parent)
        #fonts stored for a consistent UI
        self.fonts = ['Helvetica 14 bold', 'Helvetica 11 bold italic', 'Helvetica 10', 'Helvetica 10 italic',
                      'Helvetica 11 bold', 'Helvetica 22 bold']
        #the following are characters that inputted names cannot have. This list is used in the validName function
        self.invalidChar = ['0','1','2','3','4','5','6','7','8','9','!','£','$','%','^','&','*','(',')','+','=','{','}',
                       '[',']',';','@','~','#','<','>','?','/','\\','|','¬','`','¦','"',"'"]

        #used for the treeview sorting algorithm
        self.lastCol = [None, 1]
        #used to generate pseudo random unique IDs
        self.IDCounter = 1

        #initialising certain variables for a user currently not logged in
        self.registrationProgress = False
        self.create_user = False
        self.user_details = {}
        self.login_status = False
        self.redirect = []
        self.metadata = {}

        #parses the config.ini file to get the mysql data, the ebay public and private key and the email addresses for
        # email functionality
        self.db_config = self.parse_file('config.ini', 'mysql')
        self.ini_db = self.db_config.copy()
        self.ebay_config = self.parse_file('config.ini', 'eBay_key')
        self.main_db = self.db_config.copy()
        self.mail_details = self.parse_file('config.ini', 'company')
        self.MailServer = (self.mail_details['server_email'],
                           "".join([x for x in self.mail_details['server_email_password'] if
                                    x not in ['*', '&', '9', '7', '3',
                                              '4', '5', '6', '8', '|',
                                              '~', '^', '+', '?', '#',
                                              'l', 's']])
                           )
        self.frames = {}
        self.currentFrame = ""
        #this initialises all the frames needed before logging in. It instantiates these classes and creates the
        # frames
        self._pendingF = [[Load, '#ffffff'], [Login, '#ffffff'], [Create, '#ffffff']]
        for F in self._pendingF:
            self.instantiateFrame(F)

        #loads the loading frame which introduces the program and starts the 'cinematic' entrance
        self.changePage("Load", 'Starting Up...')

        #calls the method 'configure' from the module paypalrestsdk and configures it with the client secret and ID
        paypal.configure(self.parse_file(filename='config.ini', section=paypal_mode))
        logging.basicConfig(level=logging.INFO)

    def registrationTicker(self, companyName=False):
        #a variable that is false when a user is not in the process of registering a company, and true when they are.
        self.registrationProgress = companyName

    def getCreateUser(self):
        #outputs to other frame classes whether a new user is being created
        return self.create_user

    def getDBdetails(self):
        #get db details for the particular company
        return self.db_config.copy()

    def getInitDB(self):
        #get the global db details which stores the data on each company, each username and the corresponding
        # company, and the general stock information of each company. It helps link the user to their corresponding
        # company and load their company's database
        return self.ini_db.copy()

    def getEbayDetails(self):
        # returns the ebay api key
        return self.ebay_config

    def output_user(self):
        # returns the user's details if a user is logged in. Otherwise, just returns false
        if self.login_status:
            return self.user_details.copy()
        else:
            return False

    def update_user(self, details):
        # updates the Main classes copy of the user's details with the argument for 'details'
        self.user_details = details.copy()

    def get_company_metadata(self):
        #returns the data stored on the company from the main gobal database
        return self.metadata.copy()

    def getRedirect(self):
        #gets the item from the top of the redirection stack
        if len(self.redirect) >= 1:
            return self.redirect.pop(-1)
        else:
            return False

    def getAppearance(self, type):
        #returns the appearance data structure requested. Currently, only returns font since that is the only data
        # consistent appearance element I have implemented, but if a future developer wants to add a colour scheme
        # element, then they could easily edit this and get the colour scheme for each frame class.
        if type == 'font':
            return self.fonts

    def changePage(self, new, head):
        # this whole method works to change the frame and remove the old frame (but still hold it in memory) so that
        # only the active class's frames is displayed at any one time.

        #checks if there is a frame open and if that frame is one of the class frames and if so, removes it from
        # the display
        if self.currentFrame and self.currentFrame in self.frames:
            self.frames[self.currentFrame].place_forget()
        #changes the root window title
        self._top.title(head)
        # checks if the frame that needs to be raised is the instance of a user defined class or the child of one
        # if it is the class it is displayed and the currentFrame variable is updated
        if new in self.frames:
            self.frames[new].place(relx=0, rely=0, relheight=1, relwidth=1)
            self.currentFrame = new
        else:
            #if it is a child, it is just raised to the front
            self.frames[new].tkraise()

    def parse_file(self, filename, section):
        #function that gets the specific configuration details for different elements of the program from the
        # config.ini file
        self._parser = ConfigParser()
        self._parser.read(filename)

        self._config = {}
        # if the specified section argument exists in the .ini file, it iterates through all the items and adds them
        # to _config as a dictionary
        if self._parser.has_section(section):
            item = self._parser.items(section)
            for x in item:
                self._config[x[0]] = x[1]

        else:
            # in case the section doesn't exist since the .ini file is designed to be end user configurable
            raise Exception('{0} not found in the {1} file'.format(section, filename))

        return self._config

    def addFrame(self, frames):
        # frames given as a list of lists (with frames and background colour)
        # iterates through the passed list of frames and adds them to the frames dictionary
        self.recentFrames = []
        for F in frames:
            if not F[0].__name__ in self.frames:
                frame_name = F[0].__name__
                frame = F[0](self._top, self, F[1])
                self.frames[frame_name] = frame
                self.recentFrames.append(frame)
        #used to call a method of one of the instantiated classes
        return self.recentFrames

    def destroyFrames(self, frames):
        # frames given as a list to be completely removed
        for F in frames:
            self.frames[F].destroy()
            del self.frames[F]

    def load_company(self, companyDB, personalData, IDheaders, loginData, LoginHeaders):
        #called after login to to setup the program for that user

        #sets the database so that program can load details for the user's  company
        self.setCompanyDatabase(companyDB)
        self.login_status = True
        #destroys all the pre-login frames as they are not needed anymore
        self.destroyFrames(['Load', 'Login', 'Create'])

        #adds everything from the global database on the user to the user_details dictionary
        self.user_details = {}
        for column, data in zip(IDheaders, personalData):
            self.user_details[column] = data
        for column, data in zip(LoginHeaders, loginData):
            self.user_details[column] = data

        #passes each list item to instantiate all the classes to get the frames which display the company's data
        self._pendingF = [[HomePage, '#ffffff'], [EventPage, '#ffffff'], [NotificationPage, '#ffffff'],
                  [EmployeePage, '#ffffff'],
                  [TransactionsPage,'#ffffff'],[BuyInventoryPage, '#ffffff'], [StocksPage, '#ffffff'],
                  [RefundsPage,'#ffffff'], [InventoryPage, '#ffffff'], [SettingsPage, '#ffffff']]
        if self.user_details['Access_Level'] <= 2:
            self._pendingF.append([CompanySettings, '#ffffff'])
        for F in self._pendingF:
            self.instantiateFrame(F)
        #calls the method to create the menu to navigate between all the frames
        self.build_menu(self._top, self.metadata['Company_Name'], self.user_details['Access_Level'])

    def instantiateFrame(self, frameDet):
        #actually takes a list : [frame, background Colour] and instantiates the frame with the specified background
        # colour
        #currently all BGs are white, but offers a quick colour change for someone trying to change the design layout
        frame_name = frameDet[0].__name__
        frame = frameDet[0](self._top, self, frameDet[1])
        self.frames[frame_name] = frame
        frame.place_forget()

    def setCompanyDatabase(self, compDB):
        # sets the company database configuration dictionary with the user's company so the right database is accessed
        self.db_config['database'] = compDB

    def set_company_metadata(self, companyName):
        #accesses the global database with the company's name and accesses all their details and stores it as a
        # distionary to be accessed by other classes later
        #Throughout the program I automatically get the headers rather than manually entering them in case the
        # database changes in the future to make it easier for s future developer
        try:
            self.conn = MySQLConnection(**self.ini_db)
            if self.conn.is_connected() != True:
                self.log_event(self.lineno(), "Database not connected contact the network admin.")
            self.c = self.conn.cursor()
            self.c.execute("SELECT * FROM companies WHERE Company_Name = %(name)s", {'name': companyName})
            self.headers = self.c.description
            self.metadata = {}
            for a, b in zip(self.headers, self.c.fetchone()):
                self.metadata[a[0]] = b

        except Error as e:
            #adds error with mysql  to debug log
            self.log_event(e, self.lineno())
        finally:
            self.conn.close()
            return self.metadata

    def setCreateUser(self, create=False):
        #user is being created
        self.create_user = create

    def setRedirect(self, sFrame, sTitle, dFrame, dTitle, function=None):
        # adds to redirection stack: [source frame, source title, destination frame, destination title, function to
        # be called when redirected back]
        self.redirect.append([sFrame, sTitle, dFrame, dTitle, function])
        # changes the frame to the specified destination
        self.changePage(dFrame, dTitle)

    def accelerators(self, root):
        # creates the global keyboard shortcut bindings
        root.bind('<Control-h>', lambda event: webb.open(website))
        root.bind('<Control-a>', lambda e: self.aboutInfo())
        root.bind('<Control-q>', lambda e: self.file_exit())

    def build_menu(self, root, title, UsrLevel):
        # creates the menu bar used to move between frames/pages
        self.menubar = tk.Menu(root)

        self.filemenu = tk.Menu(self.menubar, tearoff=0)
        self.filemenu.add_command(label="Home", command=lambda: self.changePage('HomePage', title + ": Home"))
        self.filemenu.add_command(label="User Settings",
                                  command=lambda: self.changePage('SettingsPage', title + ": User "
                                                                                          "Settings"))
        self.filemenu.add_separator()
        self.filemenu.add_command(label="Website", command=lambda: webb.open(website), accelerator='Ctrl+H')
        self.filemenu.add_command(label="About", command=lambda: self.aboutInfo(), accelerator='Ctrl+A')
        self.filemenu.add_separator()
        self.filemenu.add_command(label="Exit", command=lambda: self.file_exit(), accelerator='Ctrl+Q')
        self.menubar.add_cascade(label="Application", menu=self.filemenu)

        self.companymenu = tk.Menu(self.menubar, tearoff=0)
        self.companymenu.add_command(label="Notifications",
                                     command=lambda: self.changePage('NotificationPage', title + ": Notifications"))
        self.companymenu.add_command(label="Employees", command=lambda: self.changePage('EmployeePage',
                                                                                        title + ": Employees"))
        self.companymenu.add_command(label="Project Management",
                                     command=lambda: self.changePage('EventPage', title + ": Project Management"))

        #access level conditional statement means any user at a level 2 or lower will not even see an option to edit
        # company settings
        if UsrLevel <= 2:
            self.companymenu.add_separator()
            self.companymenu.add_command(label="Company Settings", command=lambda: self.setRedirect('HomePage',
                                                                                                    title + ": Home",
                                                                                                    'CompanySettings',
                                                                                                    title + ': Company Settings'))
        self.menubar.add_cascade(label="Company", menu=self.companymenu)

        self.financemenu = tk.Menu(self.menubar, tearoff=0)
        self.financemenu.add_command(label="Stocks", command=lambda: self.changePage('StocksPage',
                                                                                     title + ": Stocks"))
        self.financemenu.add_command(label="Transactions", command=lambda: self.changePage('TransactionsPage',
                                                                                        title + ": Transactions"))
        self.menubar.add_cascade(label="Finance", menu=self.financemenu)

        self.inventorymenu = tk.Menu(self.menubar, tearoff=0)
        self.inventorymenu.add_command(label="Buying", command=lambda: self.changePage('BuyInventoryPage',
                                                                                       title + ": Buying"))
        self.inventorymenu.add_command(label="Sales",
                                       command=lambda: self.changePage('RefundsPage', title + ": Sales"))
        self.inventorymenu.add_command(label="Graphs and Data", command=lambda: self.changePage('InventoryPage',
                                                                                                title + ": Inventory "
                                                                                                        "Data"))
        self.menubar.add_cascade(label="Inventory", menu=self.inventorymenu)

        #adds the menubar created in the function to the Tk() instance - root
        root.config(menu=self.menubar)

    def file_exit(self):
        #function to confirm quitting the program
        if messagebox.askokcancel('Quit', 'Really Quit?'):
            if self.registrationProgress:
                #
                try:
                    self.conn = MySQLConnection(**self.ini_db)
                    if self.conn.is_connected() != True:
                        self.log_event(self.lineno(),"Database not connected contact the network admin.")
                        return
                    self.c = self.conn.cursor()
                    self.c.execute("DELETE FROM companies WHERE Company_Name=%(name)s", {'name':self.registrationProgress})
                except Error as e:
                    self.log_event(e, self.lineno())
                finally:
                    self.conn.commit()
                    self.conn.close()

            self._top.quit()

    def isfloat(self, value):
        #checks if a value is a float
        try:
            float(value)
            return True
        except ValueError:
            return False

    def validName(self, name):
        for i in self.invalidChar:
            if i in name:
                return False
        return True

    def sortItem(self, column, tree, alternateSort=True):
        # sorts a treeview widget by the column that is passed on function call. This function is called throughout
        # the program from different classes
        self.data = []
        money = False
        #gets all tree rows
        #converts data to integer, float or (keep as) string to ensure accurate sorting of data
        for k in tree.get_children(''):
            item = tree.set(k, column)
            if not money:
                if list(item)[0] == '£':
                    money = True
            if money:
                item = list(item)
                del item[0]
                item = ''.join(item)
            if item.isnumeric():
                self.data.append([int(item), k])
            elif self.isfloat(item):
                self.data.append([float(item), k])
            else:
                self.data.append([item, k])

        #checks if the column was sorted before. This small algorithm causes the treeview list to be sorted in
        # reverse every even number of times the column header is clicked
        if self.lastCol[0] == column and alternateSort:
            if self.lastCol[1] % 2 == 0:
                self.data.sort(reverse=False, key=lambda v: v[0])
            else:
                self.data.sort(reverse=True, key=lambda v: v[0])
            self.lastCol[1] += 1
        else:
            self.data.sort(reverse=False, key=lambda v: v[0])
            self.lastCol = [column, 1]

        if money:
            for x in self.data:
                x[0] = '£'+str(x[0])

        #moves the items back into the treeview in the sorted order
        for index, (val, k) in enumerate(self.data):
            tree.move(k, '', index)

    def searchItem(self, column, word, tree, treelist):
        #searches through the treeview and removes all the items that do not contain the specified keyword
        word = word.lower()
        self.data = [(tree.set(k, column), k) for k in treelist]
        tree.detach(*treelist)
        self.index = 0
        for (val, k) in self.data:
            if word in val.lower():
                tree.detach(k)
                tree.move(k, '', self.index)
                self.index += 1

    def getID(self):
        #creates a psuedo-random id based on the date and the number of function calls (looped between 1 and 20)
        date = dt.datetime.now()
        uniqueIDlist = [str(date.year), str(date.month), str(date.day), str(date.hour), str(date.minute),
                        str(date.second), "%02d" % (self.IDCounter,)]
        id = "".join(uniqueIDlist)
        #this value added to the end stops data collisions where multiple users might be adding a piece of data at
        # the same time
        if self.IDCounter < 20:
            self.IDCounter += 1
        else:
            self.IDCounter = 1

        return id

    def email_compiler(self, *mailList):
        #each mailList item should be a list in the form: [From, To, Subject, Body, Cc]

        #establishes an SMTP connection to the gmail smtp servers on the port 587
        self.server = SMTP('smtp.gmail.com', 587)
        self.server.ehlo()
        self.server.starttls()
        # uses the configuration details taken from the config.ini file, stored in the MailServer dictionary,
        # to login to the gmail mail server to send emails from that account
        self.server.login(self.MailServer[0], self.MailServer[1])

        #sends multiple emails on the same connection
        for x in range(len(mailList)):
            #constructs the email
            self.msg = MIMEMultipart()
            self.msg['From'] = "Vortex Enterprise International"
            self.msg['To'] = mailList[x][1]
            # if a 5th item is provided for the list item, it adds this list of names to be CC'd
            if len(mailList[x]) == 5:
                self.msg['Cc'] = ",".join(mailList[x][4])
            self.msg['Subject'] = mailList[x][2]
            #also sends the details of the user that performed this action. This helps not only prevent abuse of the
            # system, but also helps the recipient know who they need to reply to
            if self.login_status is True:
                self.user_info = "Username: {0}\nFirst Name: {1}\nLast Name: {2}\nEmail: {3}\nRole: {4}" \
                    .format(self.user_details['Username'], self.user_details['First_Name'],
                            self.user_details['Last_Name'], self.user_details['Email'], self.user_details['Role'])
            else:
                self.user_info = 'Unknown - Sender did not login'
            #the message skeleton
            self.body = """
Please DO NOT REPLY to this email. The sender is: {0}
************************************************************************************        
{1}
    
    
----Details----        
{2}        
    """.format(mailList[x][0], mailList[x][3], self.user_info)

            #creates the email
            self.msg.attach(MIMEText(self.body, 'plain'))
            self.text = self.msg.as_string()
            #sends the email (with or without the CC depending on arguments)
            try:
                if len(mailList[x]) == 5:
                    self.server.sendmail(self.MailServer[0], [mailList[x][1]]+mailList[x][4], self.text)
                else:
                    self.server.sendmail(self.MailServer[0], mailList[x][1], self.text)
            except:
                tk.messagebox.showerror("Error", "That email is invalid. The program will now exit. Use the access "
                                                 "code to create a new user when next opening the application.")
                self._top.quit()
        #closes server
        self.server.quit()

    def aboutInfo(self):
        #display info about the program
        messagebox.showinfo("About Genesis™ BM", """Genesis™ Business Management
        Version: {}
        © Adrian Soosaipillai 2020
        """.format(version))

    def lineno(self):
        # Returns current line number in program
        return inspect.currentframe().f_back.f_lineno

    def log_event(self, event, line="Unknown"):
        #add text and the line number (if passed) to the debug log
        self.debug_event = [dt.datetime.now(), line, event]
        self.debug_string = "{0}:   {1}  --  Line: {2}".format(dt.datetime.now(), event, line)
        print(self.debug_string)

class Load(tk.Frame):
    def __init__(self, parent, controller, _bg):
        #initialising the frame with some key variables including the root (Tk()), controller ('Main' class), and fonts
        super().__init__(parent, bg=_bg)
        self.place(relx=0, rely=0, relheight=1, relwidth=1)
        self._top = parent
        self._controller = controller
        self.fonts = controller.getAppearance('font')

        # creating all the label widgets for the booting up page
        self._title = tk.Label(self, text="Genesis BM", fg="black", font=self.fonts[0], bg=_bg)

        self._details = tk.Label(self,
                                 text="Business Management Solution\n\nDeveloped for: Young Enterprise\n Created by: "
                                    "Adrian",font=self.fonts[1], bg=_bg)

        #adding the logo image
        self.IMG = tk.PhotoImage(file='Images/young enterprise 2.gif')
        self._Logo = tk.Label(self, image=self.IMG, bg=_bg)
        self._Logo.image = self.IMG
        #creating the button widgets for the booting up page. These widgets navigate to a registration page for a new
        # user or a log in page for an existing user
        self._login = tk.Button(self, bg=_bg, text="Login to your company", command=lambda: self._controller.changePage(
            "Login","Login"),fg='black', font=self.fonts[1])

        self._new = tk.Button(self, text="Register a new company", command=lambda: self._controller.changePage(
            "Create","Create New Company"), bg='white', fg='black', font=self.fonts[1])

        #placing all widgets onto the frame
        self._title.place(relx=0.5, rely=0.1, anchor='center')
        self._details.place(relx=0.5, rely=0.25, anchor='center')
        self._Logo.place(relx=0.5, rely=0.55, anchor='center')
        self._login.place(relx=0.3, rely=0.75, anchor='center', height=50)
        self._new.place(relx=0.7, rely=0.75, anchor='center', height=50)

class Login(tk.Frame):
    def __init__(self, parent, controller, _bg):
        # create the login frame - setting values for the root window, controlling class (Main), default
        # background colour, fonts and database configuration details
        super().__init__(parent, bg=_bg)
        self.place(x=0, y=0, relheight=1, relwidth=1)
        self._top = parent
        self._controller = controller
        self.fonts = controller.getAppearance('font')
        self.db = self._controller.getDBdetails()
        self.ini_db = self._controller.getInitDB()
        self._bg = _bg

        # variables to stored the username and password entries
        self.Username = tk.StringVar()
        self.Password = tk.StringVar()
        #counter to stop brute force
        self.attempts = 0

        #creates the widgets that make up the frame including the title, entry elements and labels instructing
        # users what to enter
        self._head = tk.Label(self, text="Welcome. Please login.", bg=_bg,
                              font=self.fonts[0])
        self._head.place(relx=0.35, rely=0.1, relheight=0.05, relwidth=0.3)
        self._usernameL = tk.Label(self, text="Username", bg=_bg, font=self.fonts[1], anchor='e')
        self._usernameL.place(relx=0.2, rely=0.2, relheight=0.05, relwidth=0.15)
        self._usernameE = tk.Entry(self, textvariable=self.Username, bg=_bg, font=self.fonts[2])
        self._usernameE.place(relx=0.38, rely=0.2, relheight=0.05, relwidth=0.32)
        self._usernameE.focus_set()

        self._passwordL = tk.Label(self, text="Password", bg=_bg, font=self.fonts[1],
                                   anchor='e')
        self._passwordL.place(relx=0.2, rely=0.28, relheight=0.05, relwidth=0.15)
        self._passwordE = tk.Entry(self, textvariable=self.Password, bg=_bg, font=self.fonts[2], show="•")
        self._passwordE.place(relx=0.38, rely=0.28, relheight=0.05, relwidth=0.32)

        #button to show and hide the password plaintext
        self._showPassB = tk.Checkbutton(self, text="Show/Hide Password", bg=_bg, activebackground=_bg,
                                         command=self.show_password,font=self.fonts[2])
        self._showPassB.place(relx=0.23, rely=0.35, relheight=0.05, relwidth=0.2)

        #binding means that it will login by just pressing enter after en
        self._passwordE.bind("<Return>", lambda e: self.check_login())
        self._usernameE.bind("<Return>", lambda e: self.check_login())

        # buttons to call the login, create user and forgotten password functions respectively
        self._loginB = tk.Button(self, text="Login", bg=_bg, command=self.check_login, font=self.fonts[2])
        self._loginB.place(relx=0.6, rely=0.35, relheight=0.05, relwidth=0.1)

        self._createUserB = tk.Button(self, text="Create User", bg=_bg, command=self.create_user, font=self.fonts[2],
                                      activebackground=_bg, highlightthickness=0, bd=0)
        self._createUserB.place(relx=0.27, rely=0.5, relheight=0.05, relwidth=0.1)

        self._forgotB = tk.Button(self, text="Forgotten Password", bg=_bg, command=self.forgotten_password,
                                  font=self.fonts[2], activebackground=_bg, highlightthickness=0, bd=0)
        self._forgotB.place(relx=0.52, rely=0.5, relheight=0.05, relwidth=0.15)

    def forgotten_password(self):
        # frame to reset password - it is called when the forgotten password button is pressed to create the frame
        # for the user to enter their email, passphrase and username so the program can reset their password
        self.ForgotFrame = tk.Frame(self, bg=self._bg)
        self.ForgotFrame.place(relx=0, rely=0, relheight=1, relwidth=1)
        self._head = tk.Label(self.ForgotFrame, text="Reset your Password", bg=self._bg, font=self.fonts[0])
        self._head.place(relx=0.35, rely=0.07, relheight=0.05, relwidth=0.3)

        self._email = tk.StringVar()
        self._emailL = tk.Label(self.ForgotFrame, text="Email", bg=self._bg, font=self.fonts[1], anchor='w')
        self._emailL.place(relx=0.3, rely=0.15, relheight=0.05, relwidth=0.13)
        self._emailE = tk.Entry(self.ForgotFrame, textvariable=self._email, bg=self._bg, font=self.fonts[2])
        self._emailE.place(relx=0.44, rely=0.15, relheight=0.05, relwidth=0.24)

        self._user = tk.StringVar()
        self._UserL = tk.Label(self.ForgotFrame, text="Username", bg=self._bg, font=self.fonts[1], anchor='w')
        self._UserL.place(relx=0.3, rely=0.22, relheight=0.05, relwidth=0.13)
        self._UserE = tk.Entry(self.ForgotFrame, textvariable=self._user, bg=self._bg, font=self.fonts[2])
        self._UserE.place(relx=0.44, rely=0.22, relheight=0.05, relwidth=0.24)

        self._passphrase = tk.StringVar()
        self._passphraseL = tk.Label(self.ForgotFrame, text="Passphrase", bg=self._bg, font=self.fonts[1], anchor='w')
        self._passphraseL.place(relx=0.3, rely=0.29, relheight=0.05, relwidth=0.13)
        self._passphraseE = tk.Entry(self.ForgotFrame, textvariable=self._passphrase, bg=self._bg, font=self.fonts[2])
        self._passphraseE.place(relx=0.44, rely=0.29, relheight=0.05, relwidth=0.24)

        #calls the function to validate all the information entered into the above three fields
        self._submitB = tk.Button(self.ForgotFrame, text="Submit", bg=self._bg, font=self.fonts[2], command=self.reset_password)
        self._submitB.place(relx=0.46, rely=0.39, relheight=0.05, relwidth=0.08)

        #when the enter button is pressed on the forgotten password frame, the program proceeds to reset the password
        self._passphraseE.bind("<Return>", lambda e: self.reset_password())

        #returns back to the main login part of the frame class
        self._backB2 = tk.Button(self.ForgotFrame, text="Back", bg=self._bg, font=self.fonts[2],
                                 command=lambda: self.back(self.ForgotFrame))
        self._backB2.place(relx=0.4, rely=0.39, relheight=0.05, relwidth=0.05)

    def reset_password(self):
        #creates a random password string
        self._randPassword = ''.join(choice(ascii_uppercase + ascii_lowercase + digits)
                                     for _ in range(10))
        # connect to general mysql database
        try:
            self.conn = MySQLConnection(**self.ini_db)
            if self.conn.is_connected() != True:
                self._controller.log_event(self._controller.lineno(),
                                           "Database not connected contact the network admin.")
                return
            self.c = self.conn.cursor()
            #get the records where the username and email match the inputs
            self.c.execute("SELECT * FROM users WHERE (Username, Email)=(%(usr)s,%(email)s)",
                           {'usr': self._user.get(), 'email': self._email.get()})
            self.company_Name = self.c.fetchall()
            #if there is a match, change the database value in the db dictionary to the user's company name
            if self.company_Name:
                self.db['database'] = self.company_Name[0][1]
            else:
                tk.messagebox.showerror("Invalid Input", "One of the entered details is incorrect.")

        except Error as e:
            self._controller.log_event(e, self._controller.lineno())
        finally:
            self.conn.close()
        #connect to the user's company's private database
        try:
            self.conn = MySQLConnection(**self.db)
            if self.conn.is_connected() != True:
                self._controller.log_event(self._controller.lineno(),
                                           "Database not connected contact the network admin.")
                return
            self.c = self.conn.cursor()
            # find a matching record for the inputted username and passphrase
            self.c.execute("SELECT * FROM login WHERE (Username, Passphrase)=(%(usr)s,%(pass)s)",
                           {'usr': self._user.get(), 'pass': self._passphrase.get()})
            if self.c.fetchall():
                #if record exists, alter password to random one created and set account status to 1, which means the
                # next time the user logs in, they will be directed to the user settings page where they can change
                # that random password to something more memorable
                self.c.execute("UPDATE login SET Password=%(passw)s, Account_Status=1 WHERE Username=%(usr)s",
                               {'passw': self._randPassword, 'usr': self._user.get()})
                #email that will be sent to the user
                self._reset_email_body = """
Dear {0},
You have recently requested a password reset for the Genesis Business Management Application. 

Your password: {1}        

If you did not authorise this, you do not need to do anything, because this email is needed to access your account. 
However, you will have to log in with this password next time regardless and change it under User Settings.
                        """.format(self._user.get(), self._randPassword)
                #send the email
                self._controller.email_compiler(["Vortex Enterprise International", self._email.get(),
                                                "Genesis BM Password Reset",
                                                self._reset_email_body])
                # confirmation of email success
                tk.messagebox.showinfo("Success",
                                       "The password has been reset successfully. Please check your email for the "
                                       "new password to login.")
                self.back(self.ForgotFrame)
            else:
                #message to user to say that an input is invalid and does not match a record
                tk.messagebox.showerror("Invalid Input", "One of the entered details is incorrect.")

        except Error as e:
            self._controller.log_event(e, self._controller.lineno())
        finally:
            self.conn.commit()
            self.conn.close()

    def create_user(self):
        #creates the frame that gets the access code input in order to create a new user
        self.CreateUsrFrame = tk.Frame(self, bg=self._bg)
        self.CreateUsrFrame.place(relx=0, rely=0, relheight=1, relwidth=1)
        #label, entry and button widgets are created and placed on the CreateUsrFrame frame
        self._head = tk.Label(self.CreateUsrFrame, text="Create New User", bg=self._bg, font=self.fonts[0])
        self._head.place(relx=0.4, rely=0.2, relheight=0.05, relwidth=0.2)
        # instructions to guide new users of the system to create their account
        self._instruct1 = tk.Label(self.CreateUsrFrame, text="Enter the access code of the company you wish to join. "
                                              "If you wish to create a new company for your user, please restart this "
                                              "application and choose 'Register a new company'.",
                                   bg=self._bg, font=self.fonts[2],
                                   wraplength=350, justify='center')
        self._instruct1.place(relx=0.28, rely=0.28, relheight=0.11, relwidth=0.44)

        self._access = tk.StringVar()
        self._accessL = tk.Label(self.CreateUsrFrame, text="Access Code", bg=self._bg, font=self.fonts[1], anchor='w')
        self._accessL.place(relx=0.3, rely=0.43, relheight=0.05, relwidth=0.13)
        self._accessE = tk.Entry(self.CreateUsrFrame, textvariable=self._access, bg=self._bg, font=self.fonts[2])
        self._accessE.place(relx=0.44, rely=0.43, relheight=0.05, relwidth=0.24)
        self._accessE.focus_set()

        self._confirmB = tk.Button(self.CreateUsrFrame, text="Go", bg=self._bg, font=self.fonts[2],
                                   command=self.add_user)
        self._confirmB.place(relx=0.49, rely=0.53, relheight=0.05, relwidth=0.05)

        self._backB1 = tk.Button(self.CreateUsrFrame, text="Back", bg=self._bg, font=self.fonts[2],
                                 command=lambda: self.back(self.CreateUsrFrame))
        self._backB1.place(relx=0.43, rely=0.53, relheight=0.05, relwidth=0.05)
        #the user doesn't have to press the 'Go' button. The bind function allows the user to just hit enter after
        # entering their access code to proceed intuitively
        self._accessE.bind("<Return>", lambda e: self.add_user())

    def back(self, frame):
        #destroys the passed frame instance and removes it from display
        frame.destroy()

    def add_user(self):
        # validation to ensure a valid access code is entered
        if not 0 < len(self._access.get()) <= 10:
            tk.messagebox.showerror('Invalid Input',
                                    'Ensure the access code does not exceed 10 characters.')
            return

        #connection to the general company database
        try:
            self.conn = MySQLConnection(**self.ini_db)
            if self.conn.is_connected() != True:
                self._controller.log_event(self._controller.lineno(),
                                           "Database not connected contact the network admin.")
                return
            self.c = self.conn.cursor()
            #get the record for the company with a matching access code
            self.c.execute("SELECT Company_Name FROM companies WHERE BINARY Access_Code=%(access)s",
                           {'access':self._access.get()})
            self._accessCompName = self.c.fetchall()
            if not self._accessCompName:
                tk.messagebox.showerror('Invalid Input','That access code is invalid.')
                return
            else:
                #if there is a match, set _accessCompName equal to the company name only - a string
                self._accessCompName = self._accessCompName[0][0]


        except Error as e:
            self._controller.log_event(e, self._controller.lineno())
        #outputs the company associated with that access code and checks if the user really wants to register for
        # that company
        if not tk.messagebox.askyesno("Create User", "Do you want create a user for the "+self._accessCompName+" "
                                                                                                               "company?"):
            self._access.set("")
            return

        #navigate to the Create Employee frame and set some of the attributes of the Main class to contain the
        # metadata of the company corresponding with the input access code
        self._controller.setCreateUser(create=True)
        self._controller.setCompanyDatabase(self._accessCompName)
        self._controller.set_company_metadata(self._accessCompName)
        self._controller.addFrame([[CreateEmployee, '#ffffff']])
        self._controller.setRedirect("Login", "Login", "CreateEmployee", "Create New User")
        #close the current create user frame
        self.back(self.CreateUsrFrame)

    def show_password(self):
        #if the password field is currently showing '•', change it to plaintext, and vice versa
        if self._passwordE.cget('show') == "":
            self._passwordE.config(show="•")
        else:
            self._passwordE.config(show="")

    def check_login(self):
        #function logs in the user after they enter the username and password inputs

        #connects to the general company database
        try:
            self.conn = MySQLConnection(**self.db)
            if self.conn.is_connected() != True:
                self._controller.log_event("Database not connected contact the network admin.",
                                           self._controller.lineno())

            self.c = self.conn.cursor()
            #gets the corresponding company name for the username input
            self.c.execute("SELECT Company_Name FROM users WHERE BINARY Username = %(user)s",
                           {'user': self.Username.get()})
            self.company_ = self.c.fetchone()


        except Error as e:
            self._controller.log_event(e, self._controller.lineno())
            self.conn.close()
            return
        finally:
            self.conn.close()

        if not self.company_:
            tk.messagebox.showerror("Incorrect Input", "Your username or password is incorrect.")
            #increment attempt counter to stop brute force attacks
            self.attempts += 1
            if self.attempts >= 3:
                self.tooManyAttempts()
            return
        #if record is present, username is valid database is changed to the corresponding company name
        if self.company_:
            self.comp_db = self.db.copy()
            self.comp_db['database'] = self.company_[0]
            #connection to private, company database
            self.conn = MySQLConnection(**self.comp_db)
            if self.conn.is_connected() != True:
                self._controller.log_event(self._controller.lineno(),
                                           "Database not connected contact the network admin.")
            self.c = self.conn.cursor()
            #checks for records matching both username and password
            self.c.execute("SELECT ID, Account_Status FROM login WHERE BINARY Username=%(user)s AND BINARY Password="
                           "%(pass)s",
                           {'user': self.Username.get(), 'pass': self.Password.get()})

            self.id = self.c.fetchone()

            if self.id:
                #if record exists and inputs are valid...

                if self.id[1] == 2:
                    #if Account Status = 2, do not allow user to login. Used to quickly block off a user's access in
                    # an emergency
                    tk.messagebox.showwarning('Account Locked', 'Your account has been locked by an admin. Please '
                                                                'contact them regarding this account ban.')
                    return
                #update the last access time to the current one for the user's record
                self.c.execute("UPDATE login SET Last_Accessed=CURRENT_TIMESTAMP WHERE Username = (%(user)s)",
                               {'user': self.Username.get()})
                #get all data about the user
                self.c.execute("SELECT * FROM staff WHERE ID=(%(id)s)",
                               {'id': self.id[0]})
                self.personalID = self.c.fetchall()[0]
                self.id_headers = self.c.description
                self.id_headers = [elem[0] for elem in self.id_headers]
                if not self.personalID:
                    self._controller.log_event("Fatal database error please contact company IT admin.",
                                               self._controller.lineno())
                    tk.messagebox.showerror("Program Error", "There is an error with the database. Please contact the "
                                                             "system admin")
                    return
                # get all data about the user's account
                self.c.execute("SELECT * FROM login WHERE ID=(%(id)s)",
                               {'id': self.id[0]})
                self.personalLogin = self.c.fetchall()[0]
                self.login_headers = self.c.description
                self.login_headers = [elem[0] for elem in self.login_headers]
                self.personalLogin = list(self.personalLogin)
                self.personalID = list(self.personalID)
                self.company = self._controller.set_company_metadata(self.company_[0])

                #output successful login
                tk.messagebox.showinfo("Success", "You have successfully logged in.")
                self.attempts = 0
                #call the load company function of the Main class to load all the company's data for the user to access
                self._controller.load_company(self.company_[0], self.personalID, self.id_headers,
                                              self.personalLogin, self.login_headers)
                self.redirection = self._controller.getRedirect()
                if not self.redirection:
                    #change dimensions of the window
                    self._top.geometry("1100x700")
                    if self.id[1] == 0:
                        # if Account Status = 0, navigate to the home page
                        self._controller.changePage("HomePage", self.company['Company_Name'])
                    elif self.id[1] == 1:
                        # if Account Status = 1, navigate to the settings page and change account status to 0
                        self._controller.changePage("SettingsPage", self.company['Company_Name']+": User Settings")
                        self.c.execute("UPDATE login SET Account_Status=0 WHERE ID=%(id)s", {'id':self.id[0]})
                else:
                    #navigate to the frame specified in the list from the redirection stack
                    self._controller.changePage(self.redirection[0], self.redirection[1])
                    self.redirection[4](self.personalLogin[0])


            else:
                tk.messagebox.showerror("Incorrect Input", "Your username or password is incorrect.")
                #increment attempt counter
                self.attempts += 1
                if self.attempts >= 3:
                    self.tooManyAttempts()

        self.conn.commit()
        self.conn.close()

    def tooManyAttempts(self):
        #if 3 failed attempts, the program will output an error message and exit the application
        tk.messagebox.showerror("Too many failed attempts", "You have performed too many failed attempts to log in. "
                                                            "The program will exit now. If you have forgotten your "
                                                            "password, please re-open the app and click "
                                                            "'Forgotten Password'.")
        self._top.quit()

class CreateEmployee(tk.Frame):
    def __init__(self, parent, controller, _bg):
        # create the 'Create Employee' page - setting values for the root window, controlling class (Main), default
        # background colour, fonts and database configuration details for both the specific company database and the
        # general 'company' database
        super().__init__(parent, bg=_bg)
        self.place(relx=0, rely=0, relheight=1, relwidth=1)
        self._top = parent
        self._controller = controller
        self.fonts = controller.getAppearance('font')
        self._bg = _bg
        self.db = controller.getDBdetails()
        self.ini_db = controller.getInitDB()
        self._currentUser = self._controller.output_user()
        #newUser values:
        # 1 = new user for existing company
        # 2 = new user for new company
        # 0 = editing existing user
        if self._controller.getCreateUser():
            self.NewUser = 1
        else:
            self.NewUser = 0
            if not self._currentUser:
                self.NewUser = 2


        #data frame is the first page of details that the user needs to fill in. The following are labels, entry and
        # combobox widgets that allow the user to input their data
        self.dataFrame = tk.Frame(self, bg=_bg)
        self.dataFrame.place(relx=0, rely=0, relheight=1, relwidth=1)

        self._head1 = tk.Label(self.dataFrame, text="User Details", bg=_bg, font=self.fonts[0])
        self._head1.place(relx=0.4, rely=0.03, relheight=0.05, relwidth=0.2)

        self.initValues()

        self._usernameL = tk.Label(self.dataFrame, text="Username", bg=_bg, font=self.fonts[1], anchor='w')
        self._usernameL.place(relx=0.3, rely=0.1, relwidth=0.15, relheight=0.05)
        self._usernameE = tk.Entry(self.dataFrame, textvariable=self._username, bg=_bg, font=self.fonts[2])
        self._usernameE.place(relx=0.45, rely=0.1, relwidth=0.2, relheight=0.05)
        self._usernameForm = tk.Label(self.dataFrame, text="Max. 15 char", bg=_bg, fg='gray', font=self.fonts[3],
                                      anchor='w')
        self._usernameForm.place(relx=0.66, rely=0.1, relheight=0.05, relwidth=0.14)

        self._FNameL = tk.Label(self.dataFrame, text="First Name", bg=_bg, font=self.fonts[1], anchor='w')
        self._FNameL.place(relx=0.3, rely=0.16, relwidth=0.15, relheight=0.05)
        self._FNameE = tk.Entry(self.dataFrame, textvariable=self._FName, bg=_bg, font=self.fonts[2])
        self._FNameE.place(relx=0.45, rely=0.16, relwidth=0.2, relheight=0.05)
        self._FNameForm = tk.Label(self.dataFrame, text="Max. 25 char", bg=_bg, fg='gray', font=self.fonts[3],
                                   anchor='w')
        self._FNameForm.place(relx=0.66, rely=0.16, relheight=0.05, relwidth=0.14)

        self._LNameL = tk.Label(self.dataFrame, text="Last Name", bg=_bg, font=self.fonts[1], anchor='w')
        self._LNameL.place(relx=0.3, rely=0.22, relwidth=0.15, relheight=0.05)
        self._LNameE = tk.Entry(self.dataFrame, textvariable=self._LName, bg=_bg, font=self.fonts[2])
        self._LNameE.place(relx=0.45, rely=0.22, relwidth=0.2, relheight=0.05)
        self._LNameForm = tk.Label(self.dataFrame, text="Max. 25 char", bg=_bg, fg='gray', font=self.fonts[3],
                                   anchor='w')
        self._LNameForm.place(relx=0.66, rely=0.22, relheight=0.05, relwidth=0.14)

        self._EmailL = tk.Label(self.dataFrame, text="Email", bg=_bg, font=self.fonts[1], anchor='w')
        self._EmailL.place(relx=0.3, rely=0.28, relwidth=0.15, relheight=0.05)
        self._EmailE = tk.Entry(self.dataFrame, textvariable=self._Email, bg=_bg, font=self.fonts[2])
        self._EmailE.place(relx=0.45, rely=0.28, relwidth=0.2, relheight=0.05)
        self._EmailForm = tk.Label(self.dataFrame, text="Max. 100 char", bg=_bg, fg='gray', font=self.fonts[3],
                                   anchor='w')
        self._EmailForm.place(relx=0.66, rely=0.28, relheight=0.05, relwidth=0.14)

        self._PhoneL = tk.Label(self.dataFrame, text="Phone Number", bg=_bg, font=self.fonts[1], anchor='w')
        self._PhoneL.place(relx=0.3, rely=0.34, relwidth=0.15, relheight=0.05)
        self._PhoneE = tk.Entry(self.dataFrame, textvariable=self._Phone, bg=_bg, font=self.fonts[2])
        self._PhoneE.place(relx=0.45, rely=0.34, relwidth=0.2, relheight=0.05)
        self._PhoneForm = tk.Label(self.dataFrame, text="Max. 12 char", bg=_bg, fg='gray', font=self.fonts[3],
                                   anchor='w')
        self._PhoneForm.place(relx=0.66, rely=0.34, relheight=0.05, relwidth=0.14)

        self._dobL = tk.Label(self.dataFrame, text="Date Of Birth", bg=_bg, font=self.fonts[1], anchor='w')
        self._dobL.place(relx=0.3, rely=0.40, relwidth=0.15, relheight=0.05)
        self._dobE1 = tk.Entry(self.dataFrame, textvariable=self._DOB[0], bg=_bg, font=self.fonts[2],justify='center')
        self._dobE1.place(relx=0.45, rely=0.4, relheight=0.05, relwidth=0.05)
        tk.Label(self.dataFrame, text="/", bg=_bg, font=self.fonts[2], anchor='w').place(relx=0.5, rely=0.4,
                                                                                         relheight=0.05,
                                                                                         relwidth=0.01)
        self._dobE2 = tk.Entry(self.dataFrame, textvariable=self._DOB[1], bg=_bg, font=self.fonts[2],justify='center')
        self._dobE2.place(relx=0.51, rely=0.4, relheight=0.05, relwidth=0.05)
        tk.Label(self.dataFrame, text="/", bg=_bg, font=self.fonts[2], anchor='w').place(relx=0.56, rely=0.4,
                                                                                         relheight=0.05,
                                                                                         relwidth=0.01)
        self._dobE3 = tk.Entry(self.dataFrame, textvariable=self._DOB[2], bg=_bg, font=self.fonts[2],justify='center')
        self._dobE3.place(relx=0.57, rely=0.4, relheight=0.05, relwidth=0.05)
        self._dobForm = tk.Label(self.dataFrame, text="DD/MM/YYYY", bg=_bg, fg='gray', font=self.fonts[3], anchor='w')
        self._dobForm.place(relx=0.66, rely=0.4, relheight=0.05, relwidth=0.14)

        self._DateEmployL = tk.Label(self.dataFrame, text="Date Employed", bg=_bg, font=self.fonts[1], anchor='w')
        self._DateEmployL.place(relx=0.3, rely=0.46, relwidth=0.15, relheight=0.05)
        self._DateEmployE1 = tk.Entry(self.dataFrame, textvariable=self._DateEmploy[0], bg=_bg, font=self.fonts[2],justify='center')
        self._DateEmployE1.place(relx=0.45, rely=0.46, relheight=0.05, relwidth=0.05)
        tk.Label(self.dataFrame, text="/", bg=_bg, font=self.fonts[2], anchor='w').place(relx=0.5, rely=0.46,
                                                                                         relheight=0.05,
                                                                                         relwidth=0.01)
        self._DateEmployE2 = tk.Entry(self.dataFrame, textvariable=self._DateEmploy[1], bg=_bg, font=self.fonts[2],justify='center')
        self._DateEmployE2.place(relx=0.51, rely=0.46, relheight=0.05, relwidth=0.05)
        tk.Label(self.dataFrame, text="/", bg=_bg, font=self.fonts[2], anchor='w').place(relx=0.56, rely=0.46,
                                                                                         relheight=0.05,
                                                                                         relwidth=0.01)
        self._DateEmployE3 = tk.Entry(self.dataFrame, textvariable=self._DateEmploy[2], bg=_bg, font=self.fonts[2],justify='center')
        self._DateEmployE3.place(relx=0.57, rely=0.46, relheight=0.05, relwidth=0.05)
        self._DateEmployForm = tk.Label(self.dataFrame, text="DD/MM/YYYY", bg=_bg, fg='gray', font=self.fonts[3],
                                        anchor='w')
        self._DateEmployForm.place(relx=0.66, rely=0.46, relheight=0.05, relwidth=0.14)
        self._DateEmploy[0].set(dt.datetime.now().day)
        self._DateEmploy[1].set(dt.datetime.now().month)
        self._DateEmploy[2].set(dt.datetime.now().year)

        self._WageL = tk.Label(self.dataFrame, text="Annual Wage", bg=_bg, font=self.fonts[1], anchor='w')
        self._WageL.place(relx=0.3, rely=0.52, relwidth=0.15, relheight=0.05)
        self._WageE = tk.Entry(self.dataFrame, textvariable=self._Wage, bg=_bg, font=self.fonts[2])
        self._WageE.place(relx=0.45, rely=0.52, relwidth=0.2, relheight=0.05)
        self._WageForm = tk.Label(self.dataFrame, text="(£)", bg=_bg, fg='gray', font=self.fonts[3], anchor='w')
        self._WageForm.place(relx=0.66, rely=0.52, relheight=0.05, relwidth=0.14)

        self._AccessL = tk.Label(self.dataFrame, text="Access Level", bg=_bg, font=self.fonts[1], anchor='w')
        self._AccessL.place(relx=0.3, rely=0.58, relwidth=0.15, relheight=0.05)
        self._AccessE = tk.Spinbox(self.dataFrame, textvariable=self._Access, from_=1, to=5, state='readonly',
                                   increment=-1)
        self._AccessE.place(relx=0.45, rely=0.58, relheight=0.05, relwidth=0.1)
        self._AccessForm = tk.Label(self.dataFrame, text="1 = highest level", bg=_bg, fg='gray', font=self.fonts[3],
                                    anchor='w')
        self._AccessForm.place(relx=0.66, rely=0.58, relheight=0.05, relwidth=0.14)

        self._jobDescriptionL = tk.Label(self.dataFrame, text="Job Description", font=self.fonts[1], bg=_bg, anchor='w')
        self._jobDescriptionL.place(relx=0.3, rely=0.64, relheight=0.05, relwidth=0.15)
        self._jdscrollbar = tk.Scrollbar(self.dataFrame, orient="vertical", troughcolor='white', bg='black', bd=0)
        self._jobDescriptionT = tk.Text(self.dataFrame, bg=_bg, font=self.fonts[2], wrap='word',
                                        yscrollcommand=self._jdscrollbar.set)
        self._jdscrollbar.config(command=self._jobDescriptionT.yview)
        self._jdscrollbar.place(relx=0.65, rely=0.64, relheight=0.15)
        self._jobDescriptionT.place(relx=0.45, rely=0.64, relheight=0.15, relwidth=0.2)

        self._nationalityL = tk.Label(self.dataFrame, text="Nationality", font=self.fonts[1], bg=_bg, anchor='w')
        self._nationalityL.place(relx=0.3, rely=0.8, relheight=0.05, relwidth=0.15)
        self._nationalityE = ttk.Combobox(self.dataFrame, textvariable=self._Nationality)
        self._nationalityE.place(relx=0.45, rely=0.8, relheight=0.05, relwidth=0.2)
        self._nationalityForm = tk.Label(self.dataFrame, text="Max. 50 char", fg='gray', font=self.fonts[3], bg=_bg,
                                         anchor='w')
        self._nationalityForm.place(relx=0.66, rely=0.8, relheight=0.05, relwidth=0.14)

        self._nationalityE.bind("<Return>", lambda e: self.next())

        #returns back to the previous major frame class that navigated to 'CreateEmployee'
        self._backB = tk.Button(self.dataFrame, text="Back", bg=_bg, font=self.fonts[2], command=self.back)
        self._backB.place(relx=0.4, rely=0.9, relheight=0.05, relwidth=0.1)

        #Proceeds to validate the inputs and if all are valid, the next set of inputs needed will be displayed
        self._nextB = tk.Button(self.dataFrame, text="Next", bg=_bg, font=self.fonts[2], command=self.next)
        self._nextB.place(relx=0.55, rely=0.9, relheight=0.05, relwidth=0.1)

        #creates the frame with the second set of inputs to create a new employee
        self.moreDataFrame()
        #if the user isn't creating a user for a newly registered company, check the company's private database to
        # get options for the combobox widgets
        if self.NewUser != 2:
            self.createOptions()
        #raise the frame with the first set of inputs to the top of the program
        self.dataFrame.tkraise()

    def initValues(self):
        #instantiates all the tkinter variable classes to store inputs made in the two frames
        self._username = tk.StringVar()
        self._FName = tk.StringVar()
        self._LName = tk.StringVar()
        self._Email = tk.StringVar()
        self._Phone = tk.StringVar()
        self._DOB = [tk.IntVar(), tk.IntVar(), tk.IntVar()]
        self._DateEmploy = [tk.IntVar(), tk.IntVar(), tk.IntVar()]
        self._Wage = tk.StringVar()
        self._Access = tk.IntVar()
        self._Nationality = tk.StringVar()
        # _jobDescriptionT
        self._Role = tk.StringVar()
        self._Section = tk.StringVar()
        self._Specialisation = tk.StringVar()
        self._Passphrase = tk.StringVar()
        #_noteT

    def fillValues(self, userDetails):
        #if a user is being edited, a call to this function will be made with the details of that user being passed
        # as 'userDetails'. This function proceeds to fill all the input widgets with the user's information
        self._username.set(userDetails['Username'])
        self._FName.set(userDetails['First_Name'])
        self._LName.set(userDetails['Last_Name'])
        self._Email.set(userDetails['Email'])
        self._Phone.set(userDetails['Phone_Number'])

        self._DOB[0].set(userDetails['Date_Of_Birth'].day)
        self._DOB[1].set(userDetails['Date_Of_Birth'].month)
        self._DOB[2].set(userDetails['Date_Of_Birth'].year)

        self._DateEmploy[0].set(userDetails['Date_Employed'].day)
        self._DateEmploy[1].set(userDetails['Date_Employed'].month)
        self._DateEmploy[2].set(userDetails['Date_Employed'].year)

        self._Wage.set(userDetails['Annual_Wage'])
        self._Access.set(userDetails['Access_Level'])
        self._Nationality.set(userDetails['Nationality'])

        self._jobDescriptionT.delete(1.0, 'end')
        self._jobDescriptionT.insert('end', userDetails['Job_Description'])

        self._Role.set(userDetails['Role'])
        self._Section.set(userDetails['Section'])
        self._Specialisation.set(userDetails['Specialisation'])
        self._Passphrase.set(userDetails['Passphrase'])

        self._noteT.delete(1.0, 'end')
        self._noteT.insert('end', userDetails['Notes'])

        #makes the userDetails argument an attribute of this class that can be accessed from other functions in this
        # class
        self.newUserDetails = userDetails

    def createOptions(self):
        #selects all the different values for nationality, role, section and specialisation stored in the company's
        # private database for the user to select
        try:
            self.conn = MySQLConnection(**self.db)
            if self.conn.is_connected() != True:
                self._controller.log_event("Database not connected contact the network admin.",
                                           self._controller.lineno())
            self.c = self.conn.cursor()
            self.c.execute("SELECT DISTINCT Nationality FROM staff")
            self._nationalityOptions = [elem[0] for elem in self.c.fetchall()]
            self._nationalityE.config(values=self._nationalityOptions)

            self.c.execute("SELECT DISTINCT Role FROM staff")
            self._roleOptions = [elem[0] for elem in self.c.fetchall()]
            self._roleE.config(values=self._roleOptions)

            self.c.execute("SELECT DISTINCT Section FROM staff")
            self._sectionOptions = [elem[0] for elem in self.c.fetchall()]
            self._sectionE.config(values=self._sectionOptions)

            self.c.execute("SELECT DISTINCT Specialisation FROM staff")
            self._specialisationOptions = [elem[0] for elem in self.c.fetchall()]
            self._specialisationE.config(values=self._specialisationOptions)

        except Error as e:
            self._controller.log_event(e, self._controller.lineno())
        finally:
            self.conn.commit()
            self.conn.close()

    def back(self, success=False):
        #gets the last inserted item of the redirect stack from Main class
        self.backRedirect = self._controller.getRedirect()
        #if the user has been successfully registered, the program will send them an email, display a success
        # message, set the create user attribute to False in the Main class and call a function if there is a 5th
        # item in the redirection list
        if success:
            self.registerEmail()
            tk.messagebox.showinfo("Success","An email has been sent containing your password so you can login.")
            self._controller.setCreateUser()
        #change the top frame of the program to the source frame item in the redirection list
        self._controller.changePage(self.backRedirect[0], self.backRedirect[1])
        #destroy the current frame instance of this class
        self._controller.destroyFrames(['CreateEmployee'])
        self._controller.setCreateUser()
        if success and self.backRedirect[4]:
            self.backRedirect[4]()


    def next(self):
        #validation for the inputs entered into the first frame of the CreateEmployee frame class
        if not 0 < len(self._username.get()) <= 15:
            tk.messagebox.showerror('Invalid Input', 'Ensure the username is less than 16 characters.')
            return
        try:
            #connects to the general 'company' database to check if the username exists in the users table
            self.conn = MySQLConnection(**self.ini_db)
            if self.conn.is_connected() != True:
                self._controller.log_event(self._controller.lineno(),
                                           "Database not connected contact the network admin.")
            self.c = self.conn.cursor()
            self.c.execute("SELECT * FROM users WHERE BINARY Username=%(user)s", {'user': self._username.get()})
            if self.c.fetchall():
                #if the username exists and the current operation is to edit an existing user, and the username
                # doesn't match their existing username, then throw an error. If the operation is to create a new
                # user and the username exists, also throw an error
                if self.NewUser == 0:
                    if self.newUserDetails['Username'] != self._username.get():
                        tk.messagebox.showerror('Invalid Input', 'This username already exists, please choose another.')
                        return
                else:
                    tk.messagebox.showerror('Invalid Input',
                                            'This username already exists, please choose another.')
                    return

        except Error as e:
            self._controller.log_event(e, self._controller.lineno())
        finally:
            self.conn.close()

        #more validation for the inputs entered into the first frame of the CreateEmployee frame class
        if not 0 < len(self._FName.get()) <= 25:
            tk.messagebox.showerror('Invalid Input', 'Ensure the first name is less than 26 characters.')
            return
        if not self._controller.validName(self._FName.get()):
            tk.messagebox.showerror('Invalid Input', 'Ensure the first name is a valid name and does not '
                                                     'contain numbers or characters except - or _')
            return
        if not 0 < len(self._LName.get()) <= 25:
            tk.messagebox.showerror('Invalid Input', 'Ensure the last name is less than 26 characters.')
            return
        if not self._controller.validName(self._LName.get()):
            tk.messagebox.showerror('Invalid Input', 'Ensure the first name is a valid name and does not '
                                                     'contain numbers or characters except - or _')
            return

        if not 0 < len(self._Email.get()) <= 100 or not '@' in self._Email.get():
            tk.messagebox.showerror('Invalid Input', 'Ensure the email address is valid and less than 101 characters.')
            return
        if not 0 < len(self._Phone.get()) <= 12 or not self._Phone.get().isnumeric():
            tk.messagebox.showerror('Invalid Input', 'Ensure the phone number is a number less than 13 digits long.')
            return
        try:
            # constructs a datetime object from the inputs for date of birth and employment date to check if the
            # input is a valid value
            self._dateOfBirthFull = dt.date(year=self._DOB[2].get(), month=self._DOB[1].get(), day=self._DOB[0].get())
            self._dateEmployedFull = dt.date(year=self._DateEmploy[2].get(), month=self._DateEmploy[1].get(),
                                             day=self._DateEmploy[0].get())

        except ValueError:
            tk.messagebox.showerror('Invalid Input',
                                    'Enter a valid date in the form DD/MM/YYYY for the date of birth and the date employed.')
            return
        if self._dateEmployedFull < self._dateOfBirthFull:
            tk.messagebox.showerror('Invalid Input',
                                    'The date employed cannot be earlier than the date of birth.')
            return
        if self._dateOfBirthFull >= dt.date.today():
            tk.messagebox.showerror('Invalid Input',
                                    'The date of birth cannot be later than the current date.')
            return


        if not 0 < len(self._Wage.get()) <= 18:
            tk.messagebox.showerror('Invalid Input',
                                    'Enter a valid wage as an integer or decimal number with total length no more '
                                    'than 18 characters.')
            return
        try:
            #checks if there is more than 1 decimal place in the input and if the decimal value is greater than 2 digits
            parts = self._Wage.get().split('.')
            if len(parts) > 2:
                raise ValueError
            elif len(parts) == 2:
                if len(parts[1]) > 2:
                    raise ValueError

            #checks if the wage is an integer or real number
            self._wageFull = float(self._Wage.get())

        except ValueError:
            tk.messagebox.showerror('Invalid Input',
                                    'Enter a valid wage as an integer or decimal number. There can be no more than 2 '
                                    'decimal places for the value of money.')
            return


        if not 0 < len(self._jobDescriptionT.get(1.0, 'end')) <= 4294967295:
            tk.messagebox.showerror('Invalid Input',
                                    'Enter something in the job description field (must not exceed 4,294,967,295 '
                                    'characters).')
            return
        if not 0 < len(self._Nationality.get()) <= 50:
            tk.messagebox.showerror('Invalid Input',
                                    'Enter a nationality that does not exceed 50 characters.')
            return

        # if the inputs pass all validation checks, raise the second frame to get more inputs about the user
        self.extraFrame.tkraise()

    def moreDataFrame(self):
        #this function creates and places the widgets for the second input frame. It consists of labels, comboboxes,
        # text boxes and entry widgets like the first input frame
        self.extraFrame = tk.Frame(self, bg=self._bg)
        self.extraFrame.place(relx=0, rely=0, relheight=1, relwidth=1)

        self._head2 = tk.Label(self.extraFrame, text="User Details", bg=self._bg, font=self.fonts[0])
        self._head2.place(relx=0.4, rely=0.03, relheight=0.05, relwidth=0.2)

        self._roleL = tk.Label(self.extraFrame, text="Role", font=self.fonts[1], bg=self._bg, anchor='w')
        self._roleL.place(relx=0.3, rely=0.1, relheight=0.05, relwidth=0.15)
        self._roleE = ttk.Combobox(self.extraFrame, textvariable=self._Role)
        self._roleE.place(relx=0.45, rely=0.1, relheight=0.05, relwidth=0.2)
        self._roleForm = tk.Label(self.extraFrame, text="Max. 25 char", fg='gray', font=self.fonts[3], bg=self._bg,
                                  anchor='w')
        self._roleForm.place(relx=0.66, rely=0.1, relheight=0.05, relwidth=0.14)

        self._sectionL = tk.Label(self.extraFrame, text="Section", font=self.fonts[1], bg=self._bg, anchor='w')
        self._sectionL.place(relx=0.3, rely=0.16, relheight=0.05, relwidth=0.15)
        self._sectionE = ttk.Combobox(self.extraFrame, textvariable=self._Section)
        self._sectionE.place(relx=0.45, rely=0.16, relheight=0.05, relwidth=0.2)
        self._sectionForm = tk.Label(self.extraFrame, text="Max. 100 char", fg='gray', font=self.fonts[3],
                                     bg=self._bg, anchor='w')
        self._sectionForm.place(relx=0.66, rely=0.16, relheight=0.05, relwidth=0.14)

        self._specialisationL = tk.Label(self.extraFrame, text="Specialisation", font=self.fonts[1], bg=self._bg,
                                         anchor='w')
        self._specialisationL.place(relx=0.3, rely=0.22, relheight=0.05, relwidth=0.15)
        self._specialisationE = ttk.Combobox(self.extraFrame, textvariable=self._Specialisation)
        self._specialisationE.place(relx=0.45, rely=0.22, relheight=0.05, relwidth=0.2)
        self._specialisationForm = tk.Label(self.extraFrame, text="Max. 100 char", fg='gray', font=self.fonts[3],
                                            bg=self._bg, anchor='w')
        self._specialisationForm.place(relx=0.66, rely=0.22, relheight=0.05, relwidth=0.14)

        self._noteL = tk.Label(self.extraFrame, text="Notes", font=self.fonts[1], bg=self._bg, anchor='w')
        self._noteL.place(relx=0.3, rely=0.3, relheight=0.05, relwidth=0.15)
        self._noteSB = tk.Scrollbar(self.extraFrame, orient="vertical", troughcolor='white', bg='black', bd=0)
        self._noteT = tk.Text(self.extraFrame, bg=self._bg, font=self.fonts[2], wrap='word',
                                        yscrollcommand=self._noteSB.set)
        self._noteSB.config(command=self._noteT.yview)
        self._noteSB.place(relx=0.65, rely=0.3, relheight=0.15)
        self._noteT.place(relx=0.45, rely=0.3, relheight=0.15, relwidth=0.2)

        self._noteT.insert('end', 'None')
        #instruction about the passphrase that the user is about to enter to help with intuitive interface
        self._passwordInstruct = tk.Label(self.extraFrame, text="Next enter a new easy to remember passphrase. The "
                                                                "passphrase is a simple word or phrase used to "
                                                                "recover your password if you ever forget it.",
                                          bg=self._bg, font=self.fonts[2], wraplength=350, justify='center')
        self._passwordInstruct.place(relx=0.3, rely=0.5, relheight=0.15, relwidth=0.44)

        self._passphraseL = tk.Label(self.extraFrame, text="Passphrase", font=self.fonts[1], bg=self._bg, anchor='w')
        self._passphraseL.place(relx=0.3, rely=0.7, relheight=0.05, relwidth=0.15)
        self._passphraseE = tk.Entry(self.extraFrame, textvariable=self._Passphrase, bg=self._bg, font=self.fonts[2])
        self._passphraseE.place(relx=0.45, rely=0.7, relheight=0.05, relwidth=0.2)
        self._passphraseForm = tk.Label(self.extraFrame, text="Max. 25 char", fg='gray', font=self.fonts[3],
                                        bg=self._bg, anchor='w')
        self._passphraseForm.place(relx=0.66, rely=0.7, relheight=0.05, relwidth=0.14)



        #back button raises the first input frame
        self._back2B = tk.Button(self.extraFrame, text="Back", bg=self._bg, font=self.fonts[2],
                                 command=self.dataFrame.tkraise)
        self._back2B.place(relx=0.4, rely=0.9, relheight=0.05, relwidth=0.1)

        #apply button validates the new inputs and
        self._applyB = tk.Button(self.extraFrame, text="Finish", bg=self._bg, font=self.fonts[2], command=self.apply)
        self._applyB.place(relx=0.55, rely=0.9, relheight=0.05, relwidth=0.1)

    def apply(self):
        #validation for all the inputs for the second frame of the user details
        if not 0 < len(self._Role.get()) <= 25:
            tk.messagebox.showerror('Invalid Input', 'Ensure the role is less than 26 characters.')
            return
        if not 0 < len(self._Section.get()) <= 100:
            tk.messagebox.showerror('Invalid Input', 'Ensure the section is less than 101 characters.')
            return
        if not 0 < len(self._Specialisation.get()) <= 100:
            tk.messagebox.showerror('Invalid Input', 'Ensure the specialisation is less than 101 characters.')
            return
        if not 0 < len(self._Passphrase.get()) <= 25:
            tk.messagebox.showerror('Invalid Input', 'Ensure the passphrase is less than 26 characters.')
            return
        if not 0 < len(self._noteT.get(1.0, 'end')) <= 4294967295 or self._noteT.get(1.0, 'end') == '\n':
            tk.messagebox.showerror('Invalid Input',
                                    'Enter something in the notes field (must not exceed 4294967295 '
                                    'characters).')
            return

        #gets the details of the company from the Main class and make it an attribute of the CreateEmployee class
        self._company_meta = self._controller.get_company_metadata()
        if not self.NewUser == 0:
            #if creating a new user, call the function createCompany()
            self.createCompany()
        else:
            #if editing an exiting user...
            if self.newUserDetails != self._username.get():
                #if the inputted value for the username is different from the original one, change the username in
                # the users table of the general 'company' database
                try:
                    self.conn = MySQLConnection(**self.ini_db)
                    if self.conn.is_connected() != True:
                        self._controller.log_event(self._controller.lineno(),
                                                   "Database not connected contact the network admin.")
                    self.c = self.conn.cursor()
                    self.c.execute("UPDATE users SET Username=%(nUser)s WHERE Username=%(oUser)s",
                                   {'nUser': self._username.get(), 'oUser': self.newUserDetails['Username']})
                    self._controller.log_event("Successfully updated main database with username",
                                               self._controller.lineno())
                except Error as e:
                    self._controller.log_event(e, self._controller.lineno())
                finally:
                    self.conn.commit()
                    self.conn.close()
            try:
                #update the staff and login table of the private company's database with the new details
                self.conn = MySQLConnection(**self.db)
                if self.conn.is_connected() != True:
                    self._controller.log_event(self._controller.lineno(),
                                               "Database not connected contact the network admin.")
                self.c = self.conn.cursor()
                self.c.execute("UPDATE staff SET First_Name=%(fn)s, Last_Name=%(ln)s, Nationality=%(n)s, "
                               "Date_Of_Birth=%(dob)s, Email=%(e)s, Phone_Number=%(pn)s, Annual_Wage=%(wage)s, "
                               "Job_Description=%(job)s, Access_Level=%(access)s, Role=%(role)s, Date_Employed=%(employ)s, "
                               "Section=%(sect)s, Specialisation=%(spec)s, Notes=%(nt)s WHERE ID=%(id)s",
                               {'fn': self._FName.get(), 'ln': self._LName.get(), 'n': self._Nationality.get(),
                                'dob': self._dateOfBirthFull, 'e': self._Email.get(), 'pn': self._Phone.get(),
                                'wage': self._Wage.get(), 'job': self._jobDescriptionT.get(1.0, tk.END),
                                'access': self._Access.get(), 'role': self._Role.get(),
                                'employ': self._dateEmployedFull, 'sect': self._Section.get(),
                                'spec': self._Specialisation.get(), 'nt':self._noteT.get(1.0, 'end'),
                                'id': self.newUserDetails['ID']})


                self.c.execute("UPDATE login SET Username=%(nUser)s, Passphrase=%(pass)s WHERE ID=%(id)s",
                               {'nUser': self._username.get(), 'pass': self._Passphrase.get(), 'id': self.newUserDetails[
                                   'ID']})
                tk.messagebox.showinfo("Success", "Successfully update user settings")
            except Error as e:
                self._controller.log_event(e, self._controller.lineno())
            finally:
                self.conn.commit()
                self.conn.close()

            #go back to the last major frame class instance
            self.back()

    def createCompany(self):
        #instanciate the dbCreator class with the TK instance, the Main class and the company name as arguments
        self.DBCreate = dbCreator(self._top, self._controller, self._company_meta['Company_Name'])
        #create a 10 digit random alphanumeric password (including case sensitivity)
        self._randPassword = ''.join(choice(ascii_uppercase + ascii_lowercase + digits)
                                     for _ in range(10))
        #get all the entered inputs from the two frames and place the data in a large dictionary
        self.newUserDetails = {'Username':self._username.get(), 'Password':self._randPassword, 'Account_Status':'1',
                               'First_Name':self._FName.get(),
                               'Last_Name':self._LName.get(),
                               'Nationality':self._Nationality.get(),
                                'Date_Of_Birth':self._dateOfBirthFull, 'Email':self._Email.get(), 'Phone_Number':self._Phone.get(),
                                'Annual_Wage':self._Wage.get(), 'Job_Description':self._jobDescriptionT.get(1.0, 'end'),
                                'Access_Level':self._Access.get(), 'Role':self._Role.get(),
                                'Date_Employed':self._dateEmployedFull, 'Section':self._Section.get(),
                                'Specialisation':self._Specialisation.get(), 'Passphrase':self._Passphrase.get(),
                               'Notes':self._noteT.get(1.0, 'end')}
        if self.NewUser == 1:
            #if creating a new user for a company that was already registered
            self.DBCreate.registerUser(self.newUserDetails)
        elif self.NewUser == 2:
            #if creating the first new user for a company that was just registered
            self.DBCreate.createDB(self.newUserDetails)
            #set the registrationProgress attribute in the Main class = False. This tells the program that the
            # registration is over
            self._controller.registrationTicker()

        #call the back function to navigate away from the current frame class
        self.back(success=True)

    def registerEmail(self):
        #function creates an email to send to newly registered user
        self._sender = "Vortex Enterprise International"
        self._receiver = self._Email.get()
        self._subject = "Genesis BM - complete registration"
        self._body = """
Registration for {0}

Just a few more steps to complete registration of your Genesis account. After that, you will have the ability to manage
your business anywhere.

Your details are:
    Username = {1}
    Password = {2}
    Passphrase = {3}
    
    Access Code = {4}

Use your username and the following password to sign into the Genesis BM interface. Once logged in, please change 
your password as soon as possible under user settings.  

If you want to add more employees to your company, you can either do so manually (if you have sufficient privileges), 
under the employees tab, or give them the access code you set up. They must enter this after clicking 'Create User' 
button of the login page so they set themselves as an employee of your company.

If you did not register, please ignore this email. Whoever registered under your email will not be able to do 
anything without access to your email account. 
        """.format(self._company_meta['Company_Name'], self._username.get(), self._randPassword,
                   self._Passphrase.get(), self._company_meta['Access_Code'])
        #sends the above email elements to the function, email_compiler, in the Main class which will send the email
        # using the SMTP protocol
        self._controller.email_compiler([self._sender, self._receiver, self._subject, self._body])

class CompanySettings(tk.Frame):
    def __init__(self, parent, controller, _bg):
        # create the 'Company Settings' page - setting values for the root window, controlling class (Main), default
        # background colour, fonts, the company's details and database configuration details for the general
        # 'company'
        # database
        super().__init__(parent, bg=_bg)
        self.place(relx=0, rely=0, relheight=1, relwidth=1)
        self._top = parent
        self._controller = controller
        self.fonts = controller.getAppearance('font')
        self._bg = _bg
        self.init_db = controller.getInitDB()
        self.companyMeta = controller.get_company_metadata()

        #this mode shows whether a new company is being created
        self.mode = "New"
        #a dictionary of the current access code and company name so the user can change their other details with the
        #  validation routines catching out the data due to it being already registered
        self.currentCompanyData = {'Access_Code':'', 'Company_Name':''}

        #the following are a series of label, entry and combobox widgets to enter data about the company details
        self._head1 = tk.Label(self, text="Company Details", bg=_bg, font=self.fonts[0])
        self._head1.place(relx=0.4, rely=0.03, relheight=0.05, relwidth=0.2)

        self._name = tk.StringVar()
        self._nameL = tk.Label(self, text="Company Name", bg=_bg, font=self.fonts[1],
                               anchor='w')
        self._nameL.place(relx=0.27, rely=0.1, relwidth=0.18, relheight=0.05)
        self._nameE = tk.Entry(self, textvariable=self._name, bg=_bg, font=self.fonts[2])
        self._nameE.place(relx=0.45, rely=0.1, relwidth=0.2, relheight=0.05)
        self._nameForm = tk.Label(self, text="Max. 50 char", bg=_bg, fg='gray', font=self.fonts[3],
                                  anchor='w')
        self._nameForm.place(relx=0.66, rely=0.1, relheight=0.05, relwidth=0.14)

        self._parentName = tk.StringVar()
        self._parentNameL = tk.Label(self, text="Parent Company", bg=_bg, font=self.fonts[1], anchor='w')
        self._parentNameL.place(relx=0.27, rely=0.16, relwidth=0.18, relheight=0.05)
        self._parentNameE = tk.Entry(self, textvariable=self._parentName, bg=_bg, font=self.fonts[2])
        self._parentNameE.place(relx=0.45, rely=0.16, relwidth=0.2, relheight=0.05)
        self._parentNameForm = tk.Label(self, text="Max. 50 char", bg=_bg, fg='gray', font=self.fonts[3],
                                        anchor='w')
        self._parentNameForm.place(relx=0.66, rely=0.16, relheight=0.05, relwidth=0.14)

        self._founder = tk.StringVar()
        self._founderL = tk.Label(self, text="Founder", bg=_bg, font=self.fonts[1], anchor='w')
        self._founderL.place(relx=0.27, rely=0.22, relwidth=0.18, relheight=0.05)
        self._founderE = tk.Entry(self, textvariable=self._founder, bg=_bg, font=self.fonts[2])
        self._founderE.place(relx=0.45, rely=0.22, relwidth=0.2, relheight=0.05)
        self._founderForm = tk.Label(self, text="Max. 50 char", bg=_bg, fg='gray', font=self.fonts[3],
                                     anchor='w')
        self._founderForm.place(relx=0.66, rely=0.22, relheight=0.05, relwidth=0.14)

        self._CEO = tk.StringVar()
        self._ceoL = tk.Label(self, text="CEO", bg=_bg, font=self.fonts[1], anchor='w')
        self._ceoL.place(relx=0.27, rely=0.28, relwidth=0.18, relheight=0.05)
        self._ceoE = tk.Entry(self, textvariable=self._CEO, bg=_bg, font=self.fonts[2])
        self._ceoE.place(relx=0.45, rely=0.28, relwidth=0.2, relheight=0.05)
        self._ceoForm = tk.Label(self, text="Max. 50 char", bg=_bg, fg='gray', font=self.fonts[3],
                                 anchor='w')
        self._ceoForm.place(relx=0.66, rely=0.28, relheight=0.05, relwidth=0.14)

        self._dateEstab = [tk.IntVar(), tk.IntVar(), tk.IntVar()]
        self._dateEstabL = tk.Label(self, text="Date Established", bg=_bg, font=self.fonts[1], anchor='w')
        self._dateEstabL.place(relx=0.27, rely=0.34, relwidth=0.18, relheight=0.05)
        self._dateEstabE1 = tk.Entry(self, textvariable=self._dateEstab[0], bg=_bg, font=self.fonts[2],
                                     justify='center')
        self._dateEstabE1.place(relx=0.45, rely=0.34, relheight=0.05, relwidth=0.05)
        tk.Label(self, text="/", bg=_bg, font=self.fonts[2], anchor='w').place(relx=0.5, rely=0.34,
                                                                                            relheight=0.05,
                                                                                            relwidth=0.01)
        self._dateEstabE2 = tk.Entry(self, textvariable=self._dateEstab[1], bg=_bg, font=self.fonts[2],
                                     justify='center')
        self._dateEstabE2.place(relx=0.51, rely=0.34, relheight=0.05, relwidth=0.05)
        tk.Label(self, text="/", bg=_bg, font=self.fonts[2], anchor='w').place(relx=0.56, rely=0.34,
                                                                                            relheight=0.05,
                                                                                            relwidth=0.01)
        self._dateEstabE3 = tk.Entry(self, textvariable=self._dateEstab[2], bg=_bg, font=self.fonts[2],
                                     justify='center')
        self._dateEstabE3.place(relx=0.57, rely=0.34, relheight=0.05, relwidth=0.05)
        self._dateEstabForm = tk.Label(self, text="DD/MM/YYYY", bg=_bg, fg='gray', font=self.fonts[3],
                                       anchor='w')
        self._dateEstabForm.place(relx=0.66, rely=0.34, relheight=0.05, relwidth=0.14)

        #sets the establishment date to the current date
        self._dateEstab[0].set(dt.datetime.now().day)
        self._dateEstab[1].set(dt.datetime.now().month)
        self._dateEstab[2].set(dt.datetime.now().year)


        self._country = tk.StringVar()
        self._countryL = tk.Label(self, text="Country", font=self.fonts[1], bg=_bg, anchor='w')
        self._countryL.place(relx=0.27, rely=0.4, relheight=0.05, relwidth=0.18)
        self._countryE = ttk.Combobox(self, textvariable=self._country)
        self._countryE.place(relx=0.45, rely=0.4, relheight=0.05, relwidth=0.2)
        self._countryForm = tk.Label(self, text="Max. 50 char", fg='gray', font=self.fonts[3], bg=_bg,
                                     anchor='w')
        self._countryForm.place(relx=0.66, rely=0.4, relheight=0.05, relwidth=0.14)

        self._category = tk.StringVar()
        self._categoryL = tk.Label(self, text="Category", font=self.fonts[1], bg=_bg, anchor='w')
        self._categoryL.place(relx=0.27, rely=0.46, relheight=0.05, relwidth=0.18)
        self._categoryE = ttk.Combobox(self, textvariable=self._category)
        self._categoryE.place(relx=0.45, rely=0.46, relheight=0.05, relwidth=0.2)
        self._categoryForm = tk.Label(self, text="Max. 50 char", fg='gray', font=self.fonts[3], bg=_bg,
                                      anchor='w')
        self._categoryForm.place(relx=0.66, rely=0.46, relheight=0.05, relwidth=0.14)

        self._subCategory = tk.StringVar()
        self._subCategoryL = tk.Label(self, text="Sub-category", font=self.fonts[1], bg=_bg, anchor='w')
        self._subCategoryL.place(relx=0.27, rely=0.52, relheight=0.05, relwidth=0.18)
        self._subCategoryE = ttk.Combobox(self, textvariable=self._subCategory)
        self._subCategoryE.place(relx=0.45, rely=0.52, relheight=0.05, relwidth=0.2)
        self._subCategoryForm = tk.Label(self, text="Max. 50 char", fg='gray', font=self.fonts[3], bg=_bg,
                                         anchor='w')
        self._subCategoryForm.place(relx=0.66, rely=0.52, relheight=0.05, relwidth=0.14)

        self._accessInstruct = tk.Label(self, text="The following access code should be distributed to "
                                                                "each of your employees so that they can create a "
                                                                "user account as a member of your company.",
                                        bg=self._bg, font=self.fonts[2], wraplength=350, justify='center')
        self._accessInstruct.place(relx=0.27, rely=0.6, relheight=0.15, relwidth=0.47)

        self._access = tk.StringVar()
        self._accessL = tk.Label(self, text="Access Code", bg=_bg, font=self.fonts[1], anchor='w')
        self._accessL.place(relx=0.27, rely=0.76, relwidth=0.18, relheight=0.05)
        self._accessE = tk.Entry(self, textvariable=self._access, bg=_bg, font=self.fonts[2])
        self._accessE.place(relx=0.45, rely=0.76, relwidth=0.2, relheight=0.05)
        self._accessForm = tk.Label(self, text="Max. 10 char", bg=_bg, fg='gray', font=self.fonts[3],
                                    anchor='w')
        self._accessForm.place(relx=0.66, rely=0.76, relheight=0.05, relwidth=0.14)

        #creates a random 10 character alphanumeric (case sensitive) access code and set it as the access code input
        self._randAccessKey = ''.join(choice(ascii_uppercase + ascii_lowercase + digits)
                                      for _ in range(10))
        self._access.set(self._randAccessKey)

        #back button goes to the previous major frame class
        self._backB = tk.Button(self, text="Back", bg=_bg, font=self.fonts[2], command=self.back)
        self._backB.place(relx=0.4, rely=0.9, relheight=0.05, relwidth=0.1)

        #next button validates the company details and applies it to the database
        self._nextB = tk.Button(self, text="Apply", bg=_bg, font=self.fonts[2], command=self.apply)
        self._nextB.place(relx=0.55, rely=0.9, relheight=0.05, relwidth=0.1)

        #binding means that when the user presses the enter button when focus is on the access entry widget,
        # the apply function is called
        self._accessE.bind("<Return>", lambda e: self.apply())
        #creates the list of options for the combobox inputs
        self.createOptions()

        #if a user is currently logged in, call the function update_metadata to populate the input fields with the
        # data for the already created company
        if controller.output_user():
            self.update_metadata()

    def back(self, success=False):
        #gets the last entered item from the redirect stack in the Main class
        self.backRedirect = self._controller.getRedirect()
        if success:
            #if the company details are successfully updated and there is a 5th item in the redirect list,
            # the program will execute it as a function
            if self.backRedirect[4] != None:
                self.backRedirect[4]()
        #change the top frame to the previous major frame class instance
        self._controller.changePage(self.backRedirect[0], self.backRedirect[1])

    def update_metadata(self):
        #setting all the fields of the company details page as the user's company's details
        self._name.set(self.companyMeta['Company_Name'])
        self._nameE.config(state='readonly')
        self.mode = "Edit"
        self._parentName.set(self.companyMeta['Parent_Company'])
        self._dateEstab[0].set(self.companyMeta['Date_Established'].day)
        self._dateEstab[1].set(self.companyMeta['Date_Established'].month)
        self._dateEstab[2].set(self.companyMeta['Date_Established'].year)
        self._country.set(self.companyMeta['Country'])
        self._founder.set(self.companyMeta['Founder'])
        self._CEO.set(self.companyMeta['CEO'])
        self._category.set(self.companyMeta['Category'])
        self._subCategory.set(self.companyMeta['Sub_Category'])
        self._access.set(self.companyMeta['Access_Code'])
        self.currentCompanyData['Company_Name'] = self.companyMeta['Company_Name']
        self.currentCompanyData['Access_Code'] = self.companyMeta['Access_Code']

    def apply(self):
        #validation of all the input fields for the company details
        if not 0 < len(self._name.get()) <= 50:
            tk.messagebox.showerror('Invalid Input', 'Ensure the company name is less than 51 characters.')
            return
        if not len(self._parentName.get()) <= 50:
            tk.messagebox.showerror('Invalid Input', 'Ensure the parent company name is less than 51 characters.')
            return
        if not 0 < len(self._founder.get()) <= 50:
            tk.messagebox.showerror('Invalid Input', 'Ensure the founder\'s name is less than 51 characters.')
            return
        if not 0 < len(self._CEO.get()) <= 50:
            tk.messagebox.showerror('Invalid Input', 'Ensure the CEO\'s name is less than 51 characters.')
            return
        try:
            #converting the establishment date to a datetime object to check if the input is a valid date
            self._dateEstabFull = dt.date(year=self._dateEstab[2].get(), month=self._dateEstab[1].get(),
                                          day=self._dateEstab[0].get())

        except ValueError:
            tk.messagebox.showerror('Invalid Input',
                                    'Enter a valid date in the form DD/MM/YYYY for the establishment date.')
            return
        if not self._controller.validName(self._founder.get()):
            tk.messagebox.showerror('Invalid Input', 'Ensure the founder\'s name is a valid name and does not contain numbers or characters except - or _')
            return
        if not self._controller.validName(self._CEO.get()):
            tk.messagebox.showerror('Invalid Input', 'Ensure the CEO\'s name is a valid name and does not contain '
                                                     'numbers or characters except - or _')
            return
        if not 0 < len(self._country.get()) <= 50:
            tk.messagebox.showerror('Invalid Input',
                                    'Enter a country name that does not exceed 50 characters.')
            return
        if not self._controller.validName(self._country.get()):
            tk.messagebox.showerror('Invalid Input', 'Ensure the country\'s name is a valid name and does not contain '
                                                     'numbers or characters except - or _')
            return
        if not 0 < len(self._category.get()) <= 50:
            tk.messagebox.showerror('Invalid Input',
                                    'Enter a category that does not exceed 50 characters.')
            return
        if not 0 < len(self._subCategory.get()) <= 50:
            tk.messagebox.showerror('Invalid Input',
                                    'Enter a category that does not exceed 50 characters.')
            return

        if not 6 < len(self._access.get()) <= 10:
            tk.messagebox.showerror('Invalid Input',
                                    'Ensure the access code does not exceed 10 characters and is greater than 6 '
                                    'characters.')
            return

        try:
            self.conn = MySQLConnection(**self.init_db)
            if self.conn.is_connected() != True:
                self._controller.log_event(self._controller.lineno(),
                                           "Database not connected contact the network admin.")
                return
            self.c = self.conn.cursor()
            if self.mode == "New" and self.currentCompanyData['Company_Name'] != self._name.get():
                #if creating a new company and the entered company name doesn't equal the previous entered company
                # name, check if the company exists in the companies table in the general 'company' database
                self.c.execute("SELECT * FROM companies WHERE BINARY Company_Name=%(name)s", {'name': self._name.get()})
                if self.c.fetchall():
                    #if the company name exists, throw an error
                    tk.messagebox.showerror("Invalid Input", "That company name already exists, please change it.")
                    return
            if self.currentCompanyData['Company_Name'] == self._name.get():
                #if the company name doesn't equal the previously entered company name, delete the previously entered
                #  company name from the companies table in the general 'company' database
                self.c.execute("DELETE FROM companies WHERE Company_Name=%(name)s",
                               {'name': self.currentCompanyData['Company_Name']})
            if self.currentCompanyData['Access_Code'] != self._access.get():
                #if the access code doesn't equal the previously entered access code, check if the access code exists
                #  from the companies table in the general 'company' database
                self.c.execute("SELECT * FROM companies WHERE BINARY Access_Code=%(access)s", {'access':
                                                                                                   self._access.get()})
                if self.c.fetchall():
                    tk.messagebox.showerror("Invalid Input", "That access code already exists, please change it.")
                    return
        except Error as e:
            self._controller.log_event(e, self._controller.lineno())
        finally:
            self.conn.commit()
            self.conn.close()

        self.DBcreate = dbCreator(self._top, self._controller, self._name.get())
        self.companyDetailDict = {'Parent_Company': self._parentName.get(), 'Date_Established': self._dateEstabFull,
                                                      'Country': self._country.get(), 'Founder': self._founder.get(),
                                                      'CEO': self._CEO.get(), 'Category': self._category.get(),
                                                      'Sub_Category': self._subCategory.get(), 'Access_Code': self._access.get(),
                                                      'Company_Name': self._name.get()}

        self.currentCompanyData['Access_Code'] = self._access.get()
        self.currentCompanyData['Company_Name'] = self._name.get()


        self.DBcreate.addCompany(self.companyDetailDict)

        self._controller.set_company_metadata(self._name.get())
        self.companyMeta = self._controller.get_company_metadata()

        tk.messagebox.showinfo("Success", "Company settings saved successfully!")
        self._controller.registrationTicker(self._name.get())

        #calls the back function to navigate to the previous major frame class instance and call any function assigned
        self.back(success=True)

    def createOptions(self):
        #select all the different values for country, category, and sub-category from the companies table in the
        # general 'company' database
        try:
            self.conn = MySQLConnection(**self.init_db)
            if self.conn.is_connected() != True:
                self._controller.log_event("Database not connected contact the network admin.",
                                           self._controller.lineno())
            self.c = self.conn.cursor()
            self.c.execute("SELECT DISTINCT Country FROM companies")
            self._countryOptions = [elem[0] for elem in self.c.fetchall()]
            self._countryE.config(values=self._countryOptions)
            self.c.execute("SELECT DISTINCT Category FROM companies")
            self._categoryOptions = [elem[0] for elem in self.c.fetchall()]
            self._categoryE.config(values=self._categoryOptions)
            self.c.execute("SELECT DISTINCT Sub_Category FROM companies")
            self._subCategoryOptions = [elem[0] for elem in self.c.fetchall()]
            self._subCategoryE.config(values=self._subCategoryOptions)

        except Error as e:
            self._controller.log_event(e, self._controller.lineno())
        finally:
            self.conn.commit()
            self.conn.close()

class Create(tk.Frame):
    def __init__(self, parent, controller, _bg):
        #create the Create frame with all its widgets, setting values for the root window, controlling class (Main),
        # default background colour, font styles and database details for the general company database
        super().__init__(parent, bg=_bg)
        self.place(x=0, y=0, relheight=1, relwidth=1)
        self._top = parent
        self._controller = controller
        self.db = self._controller.getDBdetails()
        self.fonts = controller.getAppearance('font')
        self._bg = _bg

        #this creates the label and button widgets for the frame
        self._head1 = tk.Label(self, text="First ... set your Company Settings", bg=_bg, font=self.fonts[0])
        self._head1.place(relx=0.25, rely=0.15, relheight=0.05, relwidth=0.5)
        #this button navigates to the Company Settings frame class through the createCompany function
        self._compSettingsB = tk.Button(self, text="Company Settings", bg=_bg, font=self.fonts[2],
                                        command=self.createCompany)
        self._compSettingsB.place(relx=0.4, rely=0.2, relheight=0.05, relwidth=0.2)

        self._head2 = tk.Label(self, text="Register Owner", bg=_bg, font=self.fonts[0])
        self._head2.place(relx=0.4, rely=0.3, relheight=0.05, relwidth=0.2)
        self._instruct1 = tk.Label(self, text="Register login details for the owner of the "
                                              "company. This person should have the highest access level of the "
                                              "system.", bg=_bg, font=self.fonts[2], wraplength=350,
                                   justify='center')
        self._instruct1.place(relx=0.28, rely=0.35, relheight=0.08, relwidth=0.44)
        #this button navigates to teh Create Employee frame class through the registerUser function
        self._registerB = tk.Button(self, text="Register User", bg=_bg, font=self.fonts[2],
                                    command=self.registerUser, state='disabled')
        self._registerB.place(relx=0.4, rely=0.45, relheight=0.05, relwidth=0.2)

    def companyCreated(self):
        #makes the register user button clickable after the company has been created and saved to the database
        self._registerB.config(state='normal')

    def createCompany(self):
        #calls functions from the Main class to instantiate the CompanySettings frame class and add to the redirect
        # stack, telling to to call the companyCreated function on redirect back
        self._controller.addFrame([[CompanySettings, '#ffffff']])
        self._controller.setRedirect('Create', 'Create New Company', 'CompanySettings', 'Create Company',
                                     self.companyCreated)

    def registerUser(self):
        #calls functions from the Main class to instantiate the CreateEmployee frame class and add to the redirect
        # stack, telling to to call the backToLogin function on redirect back
        self._controller.addFrame([[CreateEmployee, '#ffffff']])
        self._controller.setRedirect('Create', 'Create New Company', 'CreateEmployee', 'Add an employee', self.backToLogin)

    def backToLogin(self):
        #displays the login frame
        self._controller.changePage("Login", "Login")

# Manage account settings for the logged in user only
class SettingsPage(tk.Frame):
    def __init__(self, parent, controller, _bg):
        #creates the user settings frame, setting values for the root window, controlling class (Main),
        # company's private database, font styles and background colour
        super().__init__(parent, bg=_bg)
        self.place(x=0, y=0, relheight=1, relwidth=1)
        self._top = parent
        self._controller = controller
        self.fonts = controller.getAppearance('font')
        self._bg = _bg
        self.db = self._controller.getDBdetails()

        #stores the user's password
        self.myPass = ""

        #creates label and entry widgets to display user's login details and allow editing
        self._head = tk.Label(self, text="Edit your settings", bg=_bg, font=self.fonts[0])
        self._head.place(relx=0.45, rely=0.05, relheight=0.05, relwidth=0.2)

        self._Email = tk.StringVar()
        self._emailL = tk.Label(self, text="Email", bg=_bg, font=self.fonts[1], anchor='w')
        self._emailL.place(relx=0.4, rely=0.15, relheight=0.05, relwidth=0.1)
        self._emailE = tk.Entry(self, textvariable=self._Email, bg=_bg, font=self.fonts[2])
        self._emailE.place(relx=0.55, rely=0.15, relheight=0.05, relwidth=0.15)

        self._passphrase = tk.StringVar()
        self._passphraseL = tk.Label(self, text="Passphrase", bg=_bg, font=self.fonts[1], anchor='w')
        self._passphraseL.place(relx=0.4, rely=0.22, relheight=0.05, relwidth=0.1)
        self._passphraseE = tk.Entry(self, textvariable=self._passphrase, bg=_bg, font=self.fonts[2])
        self._passphraseE.place(relx=0.55, rely=0.22, relheight=0.05, relwidth=0.15)

        self._lastAccessedL = tk.Label(self, text="Last Accessed: ", bg=_bg, font=self.fonts[1], anchor='w')
        self._lastAccessedL.place(relx=0.4, rely=0.29, relheight=0.05, relwidth=0.3)
        self._lastAccessedE = tk.Label(self, text="", bg=_bg, font=self.fonts[2], anchor='w')
        self._lastAccessedE.place(relx=0.55, rely=0.29, relheight=0.05, relwidth=0.15)

        self._accountStatus = tk.Label(self, text="Account Status: ", bg=_bg, font=self.fonts[1], anchor='w')
        self._accountStatus.place(relx=0.4, rely=0.36, relheight=0.05, relwidth=0.3)
        self._accountStatusE = tk.Label(self, text="", bg=_bg, font=self.fonts[2], anchor='w')
        self._accountStatusE.place(relx=0.55, rely=0.36, relheight=0.05, relwidth=0.15)

        self._ID = tk.Label(self, text="ID: ", bg=_bg, font=self.fonts[1], anchor='w')
        self._ID.place(relx=0.4, rely=0.42, relheight=0.05, relwidth=0.3)
        self._idE = tk.Label(self, text="", bg=_bg, font=self.fonts[2], anchor='w')
        self._idE.place(relx=0.55, rely=0.42, relheight=0.05, relwidth=0.15)

        #adds key binding to allow save data by pressing the Return key
        self._passphraseE.bind("<Return>", lambda e: self.saveChanges())
        self._emailE.bind("<Return>", lambda e: self.saveChanges())

        # Apply button to save all changes
        self._apply = tk.Button(self, text="Apply", bg=_bg, font=self.fonts[2], command=self.saveChanges)
        self._apply.place(relx=0.5, rely=0.85, relheight=0.05, relwidth=0.1)

        # Frame containing entries to change the user password
        self.PasswordFrame = tk.LabelFrame(self, text="Password", bg=_bg, bd=3)
        self.PasswordFrame.place(relx=0.4, rely=0.49, relheight=0.31, relwidth=0.3)

        #series of entry widgets, adjacent to label widgets, that show '*' instead of characters when typed in initially
        self._currentPass = tk.StringVar()
        self._currentPassL = tk.Label(self.PasswordFrame, text="Current Password", bg=_bg, font=self.fonts[2],
                                      anchor='e')
        self._currentPassL.place(relx=0.02, rely=0.05, relheight=0.12, relwidth=0.43)
        self._currentPassE = tk.Entry(self.PasswordFrame, textvariable=self._currentPass, bg=_bg, font=self.fonts[2],
                                      show="*")
        self._currentPassE.place(relx=0.5, rely=0.05, relheight=0.12, relwidth=0.4)

        self._newPass = tk.StringVar()
        self._newPassL = tk.Label(self.PasswordFrame, text="New Password", bg=_bg, font=self.fonts[2],
                                  anchor='e')
        self._newPassL.place(relx=0.02, rely=0.25, relheight=0.12, relwidth=0.43)
        self._newPassE = tk.Entry(self.PasswordFrame, textvariable=self._newPass, bg=_bg, font=self.fonts[2],
                                  show="*")
        self._newPassE.place(relx=0.5, rely=0.25, relheight=0.12, relwidth=0.4)

        self._confirmPass = tk.StringVar()
        self._confirmPassL = tk.Label(self.PasswordFrame, text="Re-Enter Password", bg=_bg, font=self.fonts[2],
                                      anchor='e')
        self._confirmPassL.place(relx=0.02, rely=0.45, relheight=0.12, relwidth=0.43)
        self._confirmPassE = tk.Entry(self.PasswordFrame, textvariable=self._confirmPass, bg=_bg, font=self.fonts[2],
                                      show="*")
        self._confirmPassE.place(relx=0.5, rely=0.45, relheight=0.12, relwidth=0.4)

        #checkbutton causes all the password fields to toggle between showing plaintext and '*' instead of the
        # characters - default is '*'
        self._showPass = tk.IntVar()
        self._showpassE = tk.Checkbutton(self.PasswordFrame, text="Show/Hide", bg=_bg, font=self.fonts[2],
                                         variable=self._showPass, command=self.showPass)
        self._showpassE.place(relx=0.02, rely=0.75, relheight=0.12, relwidth=0.4)

        #binds all the password entry widgets so when the enter button is pressed, the program calls the changePass
        # function to save the new password
        self._currentPassE.bind("<Return>", lambda e: self.changePass())
        self._newPassE.bind("<Return>", lambda e: self.changePass())
        self._confirmPassE.bind("<Return>", lambda e: self.changePass())
        #button to call the changePass function to save the new password
        self._savePass = tk.Button(self.PasswordFrame, text="Save Changes", bg=_bg, font=self.fonts[2],
                                   command=self.changePass)
        self._savePass.place(relx=0.6, rely=0.75, relheight=0.12, relwidth=0.3)
        #function called to populate the fields with user's data
        self.updateFields()

    def showPass(self):
        #toggles password fields between plaintext and '*'
        if self._showPass.get():
            #if the checkbutton is checked, convert all password fields to plaintext
            for p in (self._currentPassE, self._newPassE, self._confirmPassE):
                p.config(show="")
        else:
            #if checkbutton is not checked, convert all password fields to only show '*'
            for p in (self._currentPassE, self._newPassE, self._confirmPassE):
                p.config(show="*")

    def changePass(self):
        #changes the password to the new one entered
        if self._currentPass.get() == self.myPass:
            #if current password input is valid...
            if self._newPass.get() == self._confirmPass.get():
                #if the new and re-entered password input are equal...
                if 6 < len(self._newPass.get()) <= 25 and not self._newPass.get().isalnum() and self._newPass.get() != \
                        self.user_details['Username']:
                    #if the new password contains symbols, not the same as the username and greater than 6 characters
                    # and no more than 25 characters...
                    try:
                        self.conn = MySQLConnection(**self.db)
                        if self.conn.is_connected() != True:
                            self._controller.log_event(self._controller.lineno(),
                                                       "Database not connected contact the network admin.")
                        self.c = self.conn.cursor()
                        #updates database with new password and then outputs a success window if there is no error
                        # from the mysql call
                        self.c.execute("UPDATE login SET Password=%(pass)s WHERE ID=%(id)s",
                                       {'pass': self._newPass.get(), 'id': self.user_details["ID"]})
                        tk.messagebox.showinfo("Success",
                                               "Password was changed successfully. We suggest you now restart "
                                               "the application.")
                        #alters the password variables to the new one
                        self.user_details['Password'] = self._newPass.get()
                        self.myPass = self._newPass.get()
                        #updates the user details in the Main() class with new password
                        self._controller.update_user(self.user_details)


                    except Error as e:
                        self._controller.log_event(e, self._controller.lineno())

                    finally:
                        self.conn.commit()
                        self.conn.close()
                        #clear the password entry fields
                        for field in (self._newPass, self._currentPass, self._confirmPass):
                            field.set("")
                else:
                    tk.messagebox.showerror("Invalid Password", "Your password should be more than 6 characters long "
                                                                "(but not longer than 25 characters), "
                                                                "should include letters, numbers and at least one "
                                                                "symbol, and not be the same as your username")
            else:
                tk.messagebox.showerror("Incorrect Entry", "Your two new passwords are not identical.")
        else:
            tk.messagebox.showerror("Incorrect Entry", "You have entered an incorrect current password. Please try "
                                                       "again.")

    def updateFields(self):
        #updates the fields with the user details
        if self.myPass == "":
            #if the classes myPass variable is empty, set it equal to the user's current password
            self.user_details = self._controller.output_user()
            self.myPass = self.user_details['Password']
        self._Email.set(self.user_details['Email'])
        self._passphrase.set(self.user_details['Passphrase'])
        self._lastAccessedE.config(text=self.user_details['Last_Accessed'])
        self._accountStatusE.config(text=self.user_details['Account_Status'])
        self._idE.config(text=self.user_details['ID'])

    def saveChanges(self):
        #validates enail and passphrase inputs
        if not 0 < len(self._Email.get()) <= 100 or not '@' in self._Email.get():
            tk.messagebox.showerror("Invalid Entry", "Please use a valid email address no more than 100 characters.")
            return
        if not 0 < len(self._passphrase.get()) <= 25:
            tk.messagebox.showerror("Invalid Entry", "Ensure the passphrase is no more than 25 characters.")
            return
        #saves email and passphrase changes to the database
        try:
            self.conn = MySQLConnection(**self.db)
            if self.conn.is_connected() != True:
                self._controller.log_event(self._controller.lineno(),
                                           "Database not connected contact the network admin.")
            self.c = self.conn.cursor()
            self.c.execute("UPDATE login SET Passphrase=%(pass)s WHERE ID=%(id)s",
                           {'pass': self._passphrase.get(), 'id': self.user_details["ID"]})
            self.c.execute("UPDATE staff SET Email=%(email)s WHERE ID=%(id)s",
                           {'email': self._Email.get(), 'id': self.user_details["ID"]})

        except Error as e:
            self._controller.log_event(e, self._controller.lineno())

        finally:
            self.conn.commit()
            self.conn.close()

        #sets the database to 'company' so the program connects to the general company database
        self.database = self.db['database']
        self.db['database'] = 'company'
        try:
            self.conn = MySQLConnection(**self.db)
            if self.conn.is_connected() != True:
                self._controller.log_event(self._controller.lineno(),
                                           "Database not connected contact the network admin.")
            self.c = self.conn.cursor()
            #updates the email to the new email in the general company database
            self.c.execute("UPDATE users SET Email=%(email)s WHERE Username=%(id)s",
                           {'email': self._Email.get(), 'id': self.user_details["Username"]})
            tk.messagebox.showinfo("Success", "Email and passphrase updated successfully. We suggest you now restart "
                                              "the application.")

        except Error as e:
            self._controller.log_event(e, self._controller.lineno())

        finally:
            self.conn.commit()
            self.conn.close()
        self.db['database'] = self.database
        #updates the email and password for the user details stored in the Main() class
        self.user_details["Email"] = self._Email.get()
        self.user_details["Passphrase"] = self._passphrase.get()
        self._controller.update_user(self.user_details)

class dbCreator:
    def __init__(self, parent, controller, company_name):
        #class is used to create the database and register a user and company. This function sets the value for the
        # root winow, controlling class (Main), the company's private database, the company's general database and
        # the company name
        self._top = parent
        self._controller = controller
        self.companyName = company_name
        self.ini_db = controller.getInitDB()
        self.db = self.ini_db.copy()
        del self.db['database']

    def addCompany(self, companyDetails):
        #function adds the company to the database

        #creates copy of company details as an attribute of the class
        self.CompDetails = companyDetails.copy()

        try:
            self.conn = MySQLConnection(**self.ini_db)
            if self.conn.is_connected() != True:
                self._controller.log_event(self._controller.lineno(),
                                           "Database not connected contact the network admin.")
                return

            self.c = self.conn.cursor()
            #inserts company details into database
            self.c.execute("INSERT INTO companies VALUES (%(name)s, %(parent)s, %(date)s, "
                           "%(country)s, %(founder)s, %(ceo)s, %(cat)s, %(subCat)s,%(access)s) ON DUPLICATE KEY "
                           "UPDATE Parent_Company=%(parent)s, Date_Established=%(date)s, Country=%(country)s, "
                           "Founder=%(founder)s, CEO=%(ceo)s, Category=%(cat)s, Sub_Category=%(subCat)s, "
                           "Access_Code=%(access)s", {'parent': self.CompDetails['Parent_Company'], 'date': self.CompDetails['Date_Established'],
                                                      'country': self.CompDetails['Country'], 'founder': self.CompDetails['Founder'],
                                                      'ceo': self.CompDetails['CEO'], 'cat': self.CompDetails['Category'],
                                                      'subCat': self.CompDetails['Sub_Category'], 'access': self.CompDetails['Access_Code'],
                                                      'name': self.CompDetails['Company_Name']})
        except Error as e:
            self._controller.log_event(e, self._controller.lineno())

        finally:
            self.conn.commit()
            self.conn.close()

    def registerUser(self, _user_details):
        #function adds user to database

        #sets user details as an attribute of the class
        self.UserDetails = _user_details
        self.db['database'] = self.companyName
        try:
            self.conn = MySQLConnection(**self.db)
            if self.conn.is_connected() != True:
                self._controller.log_event(self._controller.lineno(),
                                           "Database not connected contact the network admin.")
                return

            self.c = self.conn.cursor()
            #inserts the user into the company's private database
            self.c.execute("INSERT INTO staff (First_Name, Last_Name, Nationality, Date_Of_Birth, Email, "
                           "Phone_Number, Annual_Wage, Job_Description, Access_Level, Role, Date_Employed, Section, "
                           "Specialisation, Notes) VALUES (%(fName)s, %(lName)s, %(nation)s, %(birth)s, %(email)s, "
                           "%(phone)s, %(wage)s, %(job)s, %(access)s, %(role)s, %(employed)s, %(section)s, "
                           "%(special)s, %(nt)s)", {'fName': self.UserDetails['First_Name'], 'lName': self.UserDetails[
                            'Last_Name'], 'nation': self.UserDetails['Nationality'], 'birth': self.UserDetails['Date_Of_Birth'], 'email': self.UserDetails['Email'],
                            'phone': self.UserDetails['Phone_Number'], 'wage': self.UserDetails['Annual_Wage'], 'job': self.UserDetails[
                            'Job_Description'], 'access': self.UserDetails['Access_Level'], 'role': self.UserDetails['Role'],
                            'employed': self.UserDetails['Date_Employed'], 'section': self.UserDetails['Section'], 'special': self.UserDetails[
                            'Specialisation'], 'nt':self.UserDetails['Notes']})
            self._lastRowID = self.c.lastrowid
            #gets the ID from the last input record, autocreated in the staff table from autoincrementing a counter,
            # to input the user into the login table with the correct ID
            self.c.execute("SELECT ID FROM staff WHERE ID=%(row)s", {'row':self._lastRowID})
            self._staffID = self.c.fetchall()[0][0]
            self.c.execute(
                "INSERT INTO login VALUES (%(u)s, %(p)s, CURRENT_TIMESTAMP, %(s)s, %(id)s, %(ph)s)",
                {'u': self.UserDetails['Username'], 'p': self.UserDetails['Password'], 's': self.UserDetails[
                    'Account_Status'], 'id':self._staffID, 'ph': self.UserDetails['Passphrase']})
        except Error as e:
            self._controller.log_event(e, self._controller.lineno())

        finally:
            self.conn.commit()
            self.conn.close()


        try:
            self.conn = MySQLConnection(**self.ini_db)
            if self.conn.is_connected() != True:
                self._controller.log_event(self._controller.lineno(),
                                           "Database not connected contact the network admin.")
                return

            self.c = self.conn.cursor()
            #adds the user to the general company database
            self.c.execute("INSERT INTO users VALUES (%(user)s, %(comp)s, %(email)s)", {'user': self.UserDetails[
                'Username'],'comp': self.companyName,'email': self.UserDetails['Email']})
        except Error as e:
            self._controller.log_event(e, self._controller.lineno())

        finally:
            self.conn.commit()
            self.conn.close()

    def createDB(self, user_details):
        #function creates a new database with all the tables to store data for a newly registered company
        self.UserDetails = user_details.copy()
        try:
            self.conn = MySQLConnection(**self.db)
            if self.conn.is_connected() != True:
                self._controller.log_event(self._controller.lineno(),
                                           "Database not connected contact the network admin.")
            self.c = self.conn.cursor()
            self.c.execute("CREATE DATABASE IF NOT EXISTS `{}`".format(self.companyName))
            self.c.execute("USE `{}`".format(self.companyName))
            self.c.execute("CREATE TABLE all_inventory ( Batch_ID varchar(16) NOT NULL, Date datetime NOT NULL, Item_ID bigint(10) NOT NULL, Original_Stock smallint(6) NOT NULL, Price_Paid float NOT NULL, Date_Removed datetime DEFAULT NULL, Queue tinyint(1) NOT NULL, PRIMARY KEY (Batch_ID))")
            self.c.execute("CREATE TABLE current_inventory ( Item_ID bigint(10) NOT NULL AUTO_INCREMENT, Batch_Stock int(6) NOT NULL, Total_Stock int(6) DEFAULT NULL, Price_Selling float NOT NULL, Last_Refresh datetime NOT NULL, Queue_length int(4) NOT NULL DEFAULT '1', PRIMARY KEY (Item_ID))")
            self.c.execute("CREATE TABLE ebay_price_history ( Price_ID varchar(16) NOT NULL, Date datetime NOT NULL, Item_ID bigint(10) NOT NULL, Price float NOT NULL, Search_Mode enum('auto','manual') NOT NULL, Current tinyint(1) NOT NULL, PRIMARY KEY (Price_ID))")
            self.c.execute("CREATE TABLE ebay_search_data ( Item_ID bigint(10) NOT NULL, Search_Term tinytext NOT NULL, Item_Condition varchar(30) NOT NULL, Listing varchar(30) NOT NULL, AllowReturns varchar(6) NOT NULL, Seller_Rating varchar(6) NOT NULL, Business_Type varchar(10) NOT NULL, Min_Price float NOT NULL, Max_Price float NOT NULL, Last_Updated datetime NOT NULL, PRIMARY KEY (Item_ID))")
            self.c.execute("CREATE TABLE events ( Event_ID bigint(20) NOT NULL AUTO_INCREMENT, Event_Name varchar(17) NOT NULL, Creation_Date datetime NOT NULL, Access_Level int(1) NOT NULL, Creator varchar(25) NOT NULL, Active tinyint(1) NOT NULL, Start_Date datetime DEFAULT NULL, Image varchar(30) DEFAULT NULL, Event_DB varchar(25) NOT NULL, Manual_Start tinyint(1) NOT NULL, Critical_Path varchar(255) NOT NULL, PRIMARY KEY (Event_ID))")
            self.c.execute("CREATE TABLE income_expenditure ( Transaction_ID varchar(16) NOT NULL, Details mediumtext NOT NULL, Description mediumtext NOT NULL, Price_Per_Piece double NOT NULL, Quantity int(11) NOT NULL DEFAULT '1', Total_Cost double NOT NULL, Type enum('income','expenditure') NOT NULL, Date datetime NOT NULL, Item_Batch_ID varchar(16) DEFAULT NULL, PRIMARY KEY (Transaction_ID))")
            self.c.execute("CREATE TABLE inventory ( Item_ID bigint(10) NOT NULL, Item_Name tinytext NOT NULL, Category tinytext NOT NULL, PRIMARY KEY (Item_ID))")
            self.c.execute("CREATE TABLE login ( Username varchar(15) NOT NULL, Password varchar(25) NOT NULL, Last_Accessed datetime NOT NULL, Account_Status int(1) NOT NULL DEFAULT '0', ID bigint(20) NOT NULL, Passphrase varchar(25) NOT NULL, PRIMARY KEY (Username), UNIQUE KEY login_ID_uindex (ID))")
            self.c.execute("CREATE TABLE notifications ( ID bigint(11) NOT NULL AUTO_INCREMENT, Subject tinytext NOT NULL, Due datetime NOT NULL, Message longtext NOT NULL, Category varchar(20) DEFAULT NULL, Urgency enum('low','medium','high') NOT NULL DEFAULT 'low', PRIMARY KEY (ID))")
            self.c.execute("CREATE TABLE offers ( Item_ID bigint(10) NOT NULL, On_Offer tinyint(1) NOT NULL DEFAULT '0', Price_Change float(7,2) DEFAULT NULL, Time_Limit datetime DEFAULT NULL, PRIMARY KEY (Item_ID))")
            self.c.execute("CREATE TABLE returns ( Return_ID varchar(17) NOT NULL, Return_Date datetime NOT NULL, Sale_ID varchar(120) NOT NULL, Reason tinytext NOT NULL, Returns_Staff_ID bigint(20) NOT NULL, PRIMARY KEY (Return_ID))")
            self.c.execute("CREATE TABLE sales ( Sale_ID varchar(20) NOT NULL, Date datetime NOT NULL, Item_Batch_ID varchar(16) NOT NULL, Price_Sold float(7,2) NOT NULL, Quantity int(11) NOT NULL, Total_Income float(8,2) DEFAULT NULL, staff_ID bigint(20) NOT NULL, Return_Quantity int(11) NOT NULL DEFAULT '0', PayPal tinytext, PayPal_Tax tinyint(1) DEFAULT NULL, PRIMARY KEY (Sale_ID))")
            self.c.execute("CREATE TABLE staff ( ID bigint(20) NOT NULL AUTO_INCREMENT, First_Name varchar(25) NOT NULL, Last_Name varchar(25) NOT NULL, Nationality varchar(50) NOT NULL, Date_Of_Birth date NOT NULL, Email varchar(100) NOT NULL, Phone_Number varchar(12) NOT NULL, Annual_Wage double NOT NULL, Job_Description longtext NOT NULL, Access_Level int(1) NOT NULL, Role varchar(25) NOT NULL, Date_Employed date NOT NULL, Date_Dismissed date DEFAULT NULL, Section varchar(100) NOT NULL, Specialisation varchar(100) NOT NULL, Notes longtext NOT NULL, PRIMARY KEY (ID))")
            self.c.execute("CREATE TABLE stock ( Share_ID bigint(10) NOT NULL AUTO_INCREMENT, Shareholder varchar(50) NOT NULL, Quantity int(7) NOT NULL, Price_Per_Share float DEFAULT NULL, Date datetime NOT NULL, PayPal tinytext, PayPal_Tax tinyint(1) DEFAULT NULL, PRIMARY KEY (Share_ID))")
        except Error as e:
            self._controller.log_event(e, self._controller.lineno())
        finally:
            self.conn.commit()
            self.conn.close()

        self.registerUser(self.UserDetails)

class HomePage(tk.Frame):
    def __init__(self, parent, controller, _bg):
        #creates home page frame and sets values for the root window, controlling class (Main), font styles and
        # background colour, company details, user details and the company's private database
        super().__init__(parent, bg=_bg)
        self.place(x=0, y=0, relheight=1, relwidth=1)
        self._top = parent
        self._controller = controller
        self.company = self._controller.get_company_metadata()
        self.fonts = controller.getAppearance('font')
        self.user = controller.output_user()
        self._bg = _bg
        self.db = self._controller.getDBdetails()

        #creates the image in the centre of the page which also doubles as a refresh button
        self.IMG1 = tk.PhotoImage(file='Images/young enterprise 2.gif')
        self.IMG1.zoom(3, 3)
        self._Logo = tk.Button(self, image=self.IMG1, bg=_bg, command=self.updateWidgets, relief='flat', bd=0)
        self._Logo.image = self.IMG1
        self._Logo.place(relx=0.4, rely=0.68, relheight=0.2, relwidth=0.2)


        #creates the label, labelframe, treeview and listbox widgets that view the key data of the company
        self._head1 = tk.Label(self, text="Genesis \nHomePage", bg=self._bg, font=self.fonts[5])
        self._head1.place(relx=0.4, rely=0.4, relheight=0.2, relwidth=0.2)

        self.StockFrame = tk.LabelFrame(self, text="Low Stock Alert", bg=_bg, labelanchor='se', bd=0, highlightthickness=0)
        self.StockFrame.place(relx=0.48, rely=0.05, relheight=0.3, relwidth=0.5)
        #shows products with a total stock of 5 or below that were sold in the last week
        self._stockHeaders = ('Item ID', 'Total Stock', 'Price Selling', 'Sales in the last week', 'Last Refresh')
        self._stockTree = ttk.Treeview(self.StockFrame, height=25, columns=self._stockHeaders, show='headings')
        for header in self._stockHeaders:
            self._stockTree.heading(header, text=header, command=lambda
                col=header: self._controller.sortItem(col, self._stockTree))
            self._stockTree.column(header, stretch='yes', width=8)
        self._stockTree.place(relx=0.01, rely=0.05, relheight=0.93, relwidth=0.98)
        #when clicking this table, the program navigates to the Buy Inventory page
        self._stockTree.bind("<Double-1>", lambda e: self._controller.changePage(
                'BuyInventoryPage', self.company['Company_Name'] + ": Buying"))


        self.SellingFrame = tk.LabelFrame(self, text="Top Selling", bg=_bg, labelanchor='ne', bd=0, highlightthickness=0)
        self.SellingFrame.place(relx=0.02, rely=0.03, relheight=0.5, relwidth=0.38)
        #shows sales from best to worst for products sold in the last week
        self._salesHeaders = ('Item ID', 'Quantity Sold (last week)','Total Income (last week)', 'Last Sale/Update')
        self._salesTree = ttk.Treeview(self.SellingFrame, height=25, columns=self._salesHeaders, show='headings')
        for header in self._salesHeaders:
            self._salesTree.heading(header, text=header, command=lambda
                col=header: self._controller.sortItem(col, self._salesTree))
            self._salesTree.column(header, stretch='yes', width=8)
        self._salesTree.place(relx=0.01, rely=0.02, relheight=0.96, relwidth=0.98)
        #when clicking this table, the program navigates to the sales/refund page
        self._salesTree.bind("<Double-1>", lambda e: self._controller.changePage(
                'RefundsPage', self.company['Company_Name'] + ": Sales"))


        self.EventsFrame = tk.LabelFrame(self, text="Active events", bg=_bg, labelanchor='sw', bd=0,
                                         highlightthickness=0)
        self.EventsFrame.place(relx=0.04, rely=0.65, relheight=0.28, relwidth=0.35)
        #shows all active events/projects
        self._eventHeaders = ('Event ID', 'Event Name', 'Access Level', 'Start Date')
        self._eventTree = ttk.Treeview(self.EventsFrame, height=25, columns=self._eventHeaders, show='headings')
        for header in self._eventHeaders:
            self._eventTree.heading(header, text=header, command=lambda
                col=header: self._controller.sortItem(col, self._eventTree))
            self._eventTree.column(header, stretch='yes', width=8)
        self._eventTree.place(relx=0.01, rely=0.02, relheight=0.96, relwidth=0.98)
        #when clicking this table, the program navigates to the project management page page
        self._eventTree.bind("<Double-1>", lambda e: self._controller.changePage(
                'EventPage', self.company['Company_Name'] + ": Project Management"))


        self.NotifFrame = tk.LabelFrame(self, text="This weeks notifications", bg=_bg, bd=0, highlightthickness=0)
        self.NotifFrame.place(relx=0.69, rely=0.43, relheight=0.5, relwidth=0.27)
        #shows all notifications due in the next 7 days
        self._NotifY = tk.Scrollbar(self.NotifFrame, orient="vertical", troughcolor="white",
                                       relief='flat', highlightthickness=0, bd=0, jump=1)
        self._notifLB = tk.Listbox(self.NotifFrame, bg=self._bg, font=self.fonts[4],
                                  activestyle='none', relief='flat', bd=0, highlightthickness=0,
                                   yscrollcommand=self._NotifY.set)
        self._NotifY.config(command=self._notifLB.yview)
        self._NotifY.place(relx=0.95, rely=0.01, relheight=0.99)
        self._notifLB.place(relx=0.01, rely=0.01, relheight=0.99, relwidth=0.94)
        #when clicking this table, the program navigates to the notification page
        self._notifLB.bind("<Double-1>", lambda e: self._controller.changePage(
                'NotificationPage', self.company['Company_Name'] + ": Notifications"))

        #calls a function to populate the tables with data from the database
        self.updateWidgets()

    def updateWidgets(self):
        #clears all 4 tables
        self._notifLB.delete(0, 'end')
        self._salesTree.delete(*self._salesTree.get_children())
        self._stockTree.delete(*self._stockTree.get_children())
        self._eventTree.delete(*self._eventTree.get_children())
        try:
            self.conn = MySQLConnection(**self.db)
            if self.conn.is_connected() != True:
                self._controller.log_event("Database not connected contact the network admin.",
                                           self._controller.lineno())
            self.c = self.conn.cursor()
            weekAhead = dt.datetime.now() + dt.timedelta(days=7)
            day = dt.datetime.now().replace(hour=0, minute=0, second=0, microsecond=0)
            #gets all the notifications from the database due from the current date to 7 days in advance in
            # descending order of due date
            self.c.execute("SELECT Due, Subject, Urgency FROM notifications WHERE DUE <= %(datePl)s AND DUE >= %("
                           "dateMi)s ORDER BY Due DESC",
                           {'datePl':weekAhead, 'dateMi':day})
            notifications = self.c.fetchall()
            for index, notif in enumerate(notifications):
                #inserts notifications into the respective table, colour coded according to urgency
                self._notifLB.insert(0, str(notif[0])+": "+str(notif[1]))
                if notif[2] == 'high':
                    self._notifLB.itemconfig(0, fg='red')
                elif notif[2] == 'medium':
                    self._notifLB.itemconfig(0, fg='#d68222')

            weekPast = dt.datetime.now() - dt.timedelta(days=7)
            #selects all items where the total stock is less than or equal to 5 and a sale was made in the last 7 days
            self.c.execute("SELECT current_inventory.Item_ID, current_inventory.Batch_Stock, "
                           "current_inventory.Total_Stock, current_inventory.Price_Selling, "
                           "current_inventory.Last_Refresh, SUM(sales.Quantity-sales.Return_Quantity) FROM "
                           "current_inventory "
                           "LEFT JOIN all_inventory ON (current_inventory.Item_ID = all_inventory.Item_ID) "
                           "LEFT JOIN sales ON (all_inventory.Batch_ID=sales.Item_Batch_ID)"
                           "WHERE (current_inventory.Batch_Stock+current_inventory.Total_Stock)<=5 AND sales.Date >= "
                           "%(day)s "
                           "GROUP BY current_inventory.Item_ID", {'day':weekPast})

            stock = self.c.fetchall()
            for k in stock:
                #stores items in low stock table
                self._stockTree.insert("", "end", values=k)
            #queries database for all active projects
            self.c.execute("SELECT Event_ID, Event_Name, Access_Level, Start_Date FROM events WHERE Active=True")
            events = self.c.fetchall()
            for k in events:
                #adds items to event table
                self._eventTree.insert("", "end", values=k)

            #queries database for all sales in the last 7 days ordered by descending total income
            self.c.execute("SELECT current_inventory.Item_ID, (sales.Quantity-sales.Return_Quantity),"
                           "SUM(sales.Total_Income), current_inventory.Last_Refresh FROM "
                           "current_inventory "
                           "LEFT JOIN all_inventory ON (current_inventory.Item_ID = all_inventory.Item_ID) "
                           "LEFT JOIN sales ON (all_inventory.Batch_ID=sales.Item_Batch_ID) "
                           "WHERE sales.Date >= %(day)s "
                            "GROUP BY current_inventory.Item_ID "
                           "ORDER BY SUM(sales.Total_Income) DESC ", {'day':weekPast})
            sales = self.c.fetchall()
            for k in sales:
                #adds item to sales table
                self._salesTree.insert("", "end", values=k)

        except Error as e:
            self._controller.log_event(e, self._controller.lineno())
        finally:
            self.conn.commit()
            self.conn.close()

class EmployeePage(tk.Frame):
    def __init__(self, parent, controller, _bg):
        #creates the employee page, setting values for the root window, controlling class (Main), company details,
        # user details, font styles, background colour, user access level and details for both the general and
        # company's private database
        super().__init__(parent, bg=_bg)
        self.place(x=0, y=0, relheight=1, relwidth=1)
        self._top = parent
        self._controller = controller
        self.company = self._controller.get_company_metadata()
        self.fonts = controller.getAppearance('font')
        self.user = controller.output_user()
        self._bg = _bg
        self.db = self._controller.getDBdetails()
        self.init_db = controller.getInitDB()

        #creates the label, entry, button and treeview widgets for the employee page
        self._head1 = tk.Label(self, text="Current Employees", font=self.fonts[0], bg=_bg)
        self._head1.place(relx=0.4, rely=0.02, relheight=0.05, relwidth=0.2)

        self._treeHeadings = [['ID', 'Forename', 'Surname', 'Access Level', 'Section', 'Email'], [1, 10, 10, 5, 40,
                                                                                                  40]]
        #creates the treeview widget (table) to display the list of employees
        self.staffTree = ttk.Treeview(self, height=25, columns=self._treeHeadings[0], show='headings')
        for x in range(0, 6):
            self.staffTree.heading(self._treeHeadings[0][x], text=self._treeHeadings[0][x], command=lambda
                col=self._treeHeadings[0][x]: self._controller.sortItem(col, self.staffTree))
            self.staffTree.column(self._treeHeadings[0][x], stretch='yes', width=self._treeHeadings[1][x])
        self.staffTree.place(relx=0.1, rely=0.1, relheight=0.7, relwidth=0.85)

        self.eventScroll = ttk.Scrollbar(self, orient="vertical", command=self.staffTree.yview)
        self.eventScroll.place(relx=0.95, rely=0.1, relheight=0.7)

        self.staffTree.configure(yscrollcommand=self.eventScroll.set)
        #key binding that displays staff details when double clicking a column in the table
        self.staffTree.bind("<Double-1>", lambda e: self.mini_staff_view())

        self.viewOldStaff = tk.IntVar()
        self._viewOld = tk.Checkbutton(self, text="View Dismissed Staff", bg=_bg, font=self.fonts[2],
                                           variable= self.viewOldStaff, activebackground=_bg, command=self.view_old)
        self._viewOld.place(relx=0.75, rely=0.8, relheight=0.05, relwidth=0.2)

        #calls a function to populate the table
        self.update_staff_list()

        self.Search_ = [tk.StringVar(), tk.StringVar()]
        self._searchL = tk.Label(self, text="Search", font=self.fonts[1], bg=_bg)
        self._searchL.place(relx=0.59, rely=0.89, relwidth=0.08, relheight=0.05)
        self._searchO = ttk.OptionMenu(self, self.Search_[0], '',*self._treeHeadings[0])
        self._searchO.place(relx=0.67, rely=0.89, relwidth=0.09, relheight=0.05)
        self.Search_[0].set(self._treeHeadings[0][0])
        self._searchE = tk.Entry(self, textvariable=self.Search_[1], bg=_bg, font=self.fonts[2])
        self._searchE.place(relx=0.78, rely=0.89, relwidth=0.13, relheight=0.05)
        self._searchB = tk.Button(self, text="Go", font=self.fonts[2], bg=_bg, command=lambda: self._controller.searchItem(
            self.Search_[0].get(), self.Search_[1].get(), self.staffTree, self.treeChild), activebackground=_bg,
                                  highlightthickness=0, bd=0)
        self._searchB.place(relx=0.91, rely=0.89, relheight=0.05, relwidth=0.05)

        #binding to run the search query when pressing enter in the search entry widget field
        self._searchE.bind("<Return>", lambda e: self._controller.searchItem(
            self.Search_[0].get(), self.Search_[1].get(), self.staffTree, self.treeChild))

        self._sideButton = tk.Button(self, text=":::", font=self.fonts[0], command=self.openSideBar, bg=_bg)
        self._sideButton.place(relx=0.01, rely=0.1, relheight=0.05, relwidth=0.04)

        #binding that calls the openSideBar function when the mouse hovers over the button
        self._sideButton.bind("<Enter>", lambda e:self.openSideBar())

        #creates the YE logo on the page
        self.IMG = tk.PhotoImage(file='Images/young enterprise 2.gif')
        self.IMG = self.IMG.subsample(2)
        self._Logo = tk.Label(self, image=self.IMG, bg=_bg)
        self._Logo.image = self.IMG
        self._Logo.place(relx=0.01, rely=0.85, relheight=0.15, relwidth=0.15)

        #function that creates the sidebar
        self.createSideBar()
        #binding that causes the side bar to be removed from the screen as soon as the mouse leaves it
        self.sidebar.bind("<Leave>", lambda e:self.sidebar.place_forget())

    def update_staff_list(self):
        #clears the table
        self.staffTree.delete(*self.staffTree.get_children())
        try:
            self.conn = MySQLConnection(**self.db)
            if self.conn.is_connected() != True:
                self._controller.log_event(self._controller.lineno(), "Database not connected contact the network "
                                                                      "admin.")
            self.c = self.conn.cursor()
            #if the view old staff checkbutton is checked, query the staff database for dismissed staff,
            # otherwise query for employed staff details
            if not self.viewOldStaff.get():
                self.c.execute("SELECT * FROM staff WHERE Date_Dismissed IS NULL")
            else:
                self.c.execute("SELECT * FROM staff WHERE "
                               "Date_Dismissed IS NOT NULL")
            self._staffHeadersList = [elem[0] for elem in self.c.description]
            self._staffHeaders = {elem[1]:elem[0] for elem in enumerate(self._staffHeadersList)}
            self._staffData = {elem[0]:list(elem) for elem in self.c.fetchall()}
            index = [self._staffHeaders['ID'], self._staffHeaders['First_Name'], self._staffHeaders['Last_Name'],
                        self._staffHeaders['Access_Level'], self._staffHeaders['Section'], self._staffHeaders['Email']]
            for staff in self._staffData:
                #insert each staff item into the table
                self.staffTree.insert("", "end", values=(self._staffData[staff][index[0]],
                                                         self._staffData[staff][index[1]],
                                                         self._staffData[staff][index[2]],
                                                         self._staffData[staff][index[3]],
                                                         self._staffData[staff][index[4]],
                                                         self._staffData[staff][index[5]]))

        except Error as e:
            self._controller.log_event(e, self._controller.lineno())
        finally:
            self.conn.close()
        #save the treeview data as an attribute to allow searching and sorting the treeview
        self.treeChild = self.staffTree.get_children('')

    def mini_staff_view(self):
        #function displays the staff's details in a window pop-up
        selectionID = self.staffTree.item(self.staffTree.selection(), 'values')[0]
        self._mini_staff_ls = []
        self._top.clipboard_clear()
        for a, b in zip(self._staffHeadersList, self._staffData[int(selectionID)]):
            self._mini_staff_ls.append("{0}: {1}\n".format(a, b))
            if a == "Email":
                #saves staff email to clipboard
                self._top.clipboard_append(b)
        tk.messagebox.showinfo("Staff ID: " + str(selectionID), "".join(self._mini_staff_ls))

    def add_staff(self):
        #function used to create user
        #sets the program to create a new user, instantiates the CreateEmployee class and adds to the redirect stack
        # telling the program to call back() on return
        self._controller.setCreateUser(create=True)
        self._controller.addFrame([[CreateEmployee, '#ffffff']])
        self._controller.setRedirect("EmployeePage", self.company['Company_Name']+": Employees", "CreateEmployee",
                                     "Create New User", self.back)

    def edit_staff(self):
        #function used to edit a member of staff
        #validation in case no user is selected
        if not self.staffTree.selection():
            tk.messagebox.showerror("Select staff", "Please select a member of staff from the table to edit their "
                                                   "details.")
            return
        #gets the id of the selected staff in the treeview
        self.selectionID = self.staffTree.item(self.staffTree.selection(), 'values')[0]
        self.selectionDetails = {}
        #gets the staff details from the company's private database
        try:
            self.conn = MySQLConnection(**self.db)
            if self.conn.is_connected() != True:
                self._controller.log_event(self._controller.lineno(), "Database not connected contact the network "
                                                                      "admin.")
            self.c = self.conn.cursor()
            self.c.execute("SELECT Username, Passphrase FROM login WHERE ID=%(id)s", {'id':self.selectionID})
            self._selectionDet = self.c.fetchall()[0]
            self.selectionDetails['Username'] = self._selectionDet[0]
            self.selectionDetails['Passphrase'] = self._selectionDet[1]

            self.c.execute("SELECT * FROM staff WHERE ID=%(id)s", {'id':self.selectionID})
            self.headers = self.c.description
            for a, b in zip(self.headers, self.c.fetchall()[0]):
                self.selectionDetails[a[0]] = b

        except Error as e:
            self._controller.log_event(e, self._controller.lineno())
        finally:
            self.conn.close()

        #instantiates the CreateEmployee class and adds to the redirect stack, telling the program to call back() on
        # return
        self._controller.addFrame([[CreateEmployee, '#ffffff']])[0].fillValues(self.selectionDetails)
        self._controller.setRedirect("EmployeePage", self.company['Company_Name']+": Employees", "CreateEmployee",
                                     "Edit current user", self.back)

    def remove_staff(self):
        #function used to fire a member of staff
        #validation routine in case no user is selected
        if not self.staffTree.selection():
            tk.messagebox.showerror("Select staff", "Please select a member of staff from the table to dismiss them.")
            return
        #gets the id of selected staff
        self.selectionID = self.staffTree.item(self.staffTree.selection(), 'values')[0]
        if tk.messagebox.askyesno("Confirm", "Are you sure you wish to dismiss this member of staff?"):
            #if confirmation given, deletes the staff details from the login table, but retains their details in the
            # staff table, updating the Date_dismissed to the current date
            try:
                self.conn = MySQLConnection(**self.db)
                if self.conn.is_connected() != True:
                    self._controller.log_event(self._controller.lineno(), "Database not connected contact the network "
                                                                          "admin.")
                self.c = self.conn.cursor()
                self.c.execute("UPDATE staff SET Date_Dismissed=CURRENT_DATE WHERE ID=%(id)s", {'id':self.selectionID})
                self.c.execute("SELECT Username FROM login WHERE ID=%(id)s", {'id':self.selectionID})
                self.selectionUser = self.c.fetchall()[0][0]
                self.c.execute("DELETE FROM login WHERE ID=%(id)s", {'id':self.selectionID})
            except Error as e:
                self._controller.log_event(e, self._controller.lineno())
            finally:
                self.conn.commit()
                self.conn.close()
            try:
                #deletes the user from the general company database to stop them logging in
                self.conn = MySQLConnection(**self.init_db)
                if self.conn.is_connected() != True:
                    self._controller.log_event(self._controller.lineno(), "Database not connected contact the network "
                                                                          "admin.")
                self.c = self.conn.cursor()
                self.c.execute("DELETE FROM users WHERE Username=%(usr)s", {'usr': self.selectionUser})
            except Error as e:
                self._controller.log_event(e, self._controller.lineno())
            finally:
                self.conn.commit()
                self.conn.close()
            #calls back() function to refresh staff list
            self.back()

    def view_old(self):
        #if the view dismissed staff checkbutton was checked, it changes the title, disables the edit, remove and
        # delete old staff button, and refreshes the treeview
        #access levels used to block the buttons if the user has a level lower (greater number) than listed in each
        # conditional statement
        if self.viewOldStaff.get():
            self._head1.config(text="Dismissed Employees")
            if self.user['Access_Level'] <= 3:
                self.editB.config(state="disabled")
            if self.user['Access_Level'] <= 2:
                self.removeB.config(state="disabled")
                self.delOldB.config(state='normal')
            self.update_staff_list()

        else:
            self._head1.config(text="Current Employees")
            if self.user['Access_Level'] <= 3:
                self.editB.config(state="normal")
            if self.user['Access_Level'] <= 2:
                self.removeB.config(state="normal")
                self.delOldB.config(state='disabled')
            self.update_staff_list()

    def back(self, frame=None):
        #destroys frames from display that are passed as a list argument and calls update_staff_list that updates the
        # data in the treeview
        if frame:
            frame.destroy()
        self.update_staff_list()

    def removeOld(self):
        #function removes the dismissed staff from the database completely
        if not tk.messagebox.askyesno("Sure", "Are you sure you want to delete all the old staff?"):
            #if no confirmation given, exits the function
            return
        try:
            #deletes the user's record from the database
            self.conn = MySQLConnection(**self.db)
            if self.conn.is_connected() != True:
                self._controller.log_event(self._controller.lineno(), "Database not connected contact the network "
                                                                      "admin.")
            self.c = self.conn.cursor()
            self.c.execute("DELETE FROM staff WHERE Date_Dismissed is not NULL")
        except Error as e:
            self._controller.log_event(e, self._controller.lineno())
        finally:
            self.conn.commit()
            self.conn.close()
        #calls the back() function which refreshes the table/treeview
        self.back()

    def createSideBar(self):
        #creates the button widgets inside the label frame widget that forms the sidebar that pops up when hovering
        # over or clicking the sidebar button (':::')
        self.sidebar = tk.LabelFrame(self, text="Menu", bg=self._bg, bd=3)
        self.sidebar.place(relx=0.01, rely=0.15, relheight=0.4, relwidth=0.2)
        self.addB = tk.Button(self.sidebar, text="Add Staff", font=self.fonts[2], command=self.add_staff)
        self.addB.place(relx=0.1, rely=0.05, relheight=0.1, relwidth=0.8)
        self.removeB = tk.Button(self.sidebar, text="Remove Staff", font=self.fonts[2], command=self.remove_staff)
        self.removeB.place(relx=0.1, rely=0.2, relheight=0.1, relwidth=0.8)
        self.editB = tk.Button(self.sidebar, text="Edit Staff", font=self.fonts[2], command=self.edit_staff)
        self.editB.place(relx=0.1, rely=0.35, relheight=0.1, relwidth=0.8)
        self.delOldB = tk.Button(self.sidebar, text="Remove Old Staff", font=self.fonts[2], state='disabled',
                                 command=self.removeOld)
        self.delOldB.place(relx=0.1, rely=0.5, relheight=0.1, relwidth=0.8)
        self.downloadB = tk.Button(self.sidebar, text="Download Staff", font=self.fonts[2], command=self.downloadDB)
        self.downloadB.place(relx=0.1, rely=0.65, relheight=0.1, relwidth=0.8)
        self.refreshB = tk.Button(self.sidebar, text="Refresh", font=self.fonts[2], command=self.update_staff_list)
        self.refreshB.place(relx=0.1, rely=0.8, relheight=0.1, relwidth=0.8)

        if self.user['Access_Level'] > 2:
            #if the user's access level is greater than 2, they cannot add or remove staff
            self.addB.config(state='disabled')
            self.removeB.config(state='disabled')
        if self.user['Access_Level'] > 3:
            #if the user's access level is greater than 3, they cannot edit or download staff data to a spreadsheet
            self.editB.config(state='disabled')
            self.downloadB.config(state='disabled')

        #removes the sidebar from display
        self.sidebar.place_forget()

    def openSideBar(self):
        #if the sidebar is displayed, remove it, otherwise display it
        if self.sidebar.winfo_ismapped():
            self.sidebar.place_forget()
        else:
            self.sidebar.place(relx=0.01, rely=0.15, relheight=0.4, relwidth=0.2)
            self.sidebar.focus_set()

    def downloadDB(self):
        #function to download staff details to an excel spreadsheet
        self.style0 = easyxf("align:horiz centre")
        self.style1 = easyxf('align:horiz center', num_format_str='DD-MMM-YY')
        self.style2 = easyxf('pattern: pattern solid, fore_color yellow; font: bold on;align:horiz center')

        self.book = Workbook()
        self.sheet = self.book.add_sheet("Staff")

        try:
            #queries all the staff data from the database
            self.conn = MySQLConnection(**self.db)
            if self.conn.is_connected() != True:
                self._controller.log_event(self._controller.lineno(), "Database not connected contact the network "
                                                                      "admin.")
            self.c = self.conn.cursor()

            self.c.execute("SELECT * FROM staff")
            self.header = [elem[0] for elem in self.c.description]
            #iterates through each column header of the staff database table and adds it to the spreadsheet
            for CI, col in enumerate(self.header):
                self.sheet.write(0, CI, col, self.style2)
            #iterates through each staff record to add it to the spreadsheet
            for r, row in enumerate(self.c.fetchall()):
                for c, col in enumerate(row):
                    if isinstance(col, dt.date):
                        self.sheet.write(r + 1, c, col, self.style1)
                    else:
                        self.sheet.write(r + 1, c, col, self.style0)
            #launch a pop-up file explorer to select a save location
            self.file = filedialog.asksaveasfilename(title="Save Staff Data", initialdir='Log/',
                                                     defaultextension=".xls")
            if self.file:
                #if a location is selected, save the file
                self.book.save(self.file)
        except Error as e:
            self._controller.log_event(e, self._controller.lineno())

        finally:
            self.conn.close()

class NotificationPage(tk.Frame):
    def __init__(self, parent, controller, _bg):
        #creates the notification page, setting a value for the root window, controlling class (Main), user details,
        # company details, font styles, background colour and details on accessing the company's private database
        super().__init__(parent, bg=_bg)
        self.place(x=0, y=0, relheight=1, relwidth=1)
        self._bg = _bg
        self._top = parent
        self._controller = controller
        self.company = controller.get_company_metadata()
        self.user = controller.output_user()
        self.db = controller.getDBdetails()
        self.fonts = controller.getAppearance('font')
        self.access = self.user['Access_Level']

        #initialising variables: no notification is being edited and the list is not categorised
        self._EditNotif = False
        self._categorised = False

        #creates the small blue tab that resizes when the window resizes and calls the function openSideBar when the
        # mouse hovers over it
        self._canvasTab = tk.Canvas(self, bg=_bg, bd=0, highlightthickness=0, relief='ridge', height=700)
        self._canvasTab.place(relx=0, rely=0, relheight=1, relwidth=0.02)
        self._canvasTab.height = self._canvasTab.winfo_reqheight()
        self._canvasTab.bind("<Configure>", lambda e: self.resize(self._canvasTab, e))
        self.rect = self.curved_rect(0, 25, 0, self._canvasTab.height, self._canvasTab, r=75, fill="#adccff")

        self._canvasTab.bind("<Enter>", lambda e:self.openSideBar())

        #calls the briefNotif function to create the listbox with the list of all the notifications
        self.briefNotif()
        #calls the createSideBar function to create the popup sidebar
        self.createSideBar()
        #binding makes the big side bar disappear when the mouse hovers over the frame with the list of all
        # notifications
        self.briefFrame.bind("<Enter>", lambda e:self.openSideBar(closeOnly=True))

        #frame that occupies the empty space until a notification is clicked and specific details are displayed
        self.fillerFrame = tk.Frame(self, bg=_bg)
        self.fillerFrame.place(relx=0.42, rely=0.1, relheight=0.83, relwidth=0.55)
        self.IMG1 = tk.PhotoImage(file='Images/young enterprise 2.gif')
        self._Logo2 = tk.Label(self.fillerFrame, image=self.IMG1, bg=_bg)
        self._Logo2.image = self.IMG1
        self._Logo2.place(relx=0.2, rely=0.3, relheight=0.4, relwidth=0.6)

        #creates the widgets for the frame showing full details of the notification
        self.fullNotif()
        #calls the function to populate the listbox of notifications
        self.updateItems()

    def initValues(self):
        #initialises the tkinter variables for the entry widgets
        self._Subject = tk.StringVar()
        self._Due = [tk.IntVar(), tk.IntVar(), tk.IntVar(), tk.IntVar(), tk.IntVar()]
        self._Category = tk.StringVar()
        self._Urgency = tk.StringVar()
        # _messageT

    def editValues(self, notifDetails):
        #called if editing a notification since it populates the create notification input widgets with stored
        # details of the notification
        self._EditNotif = notifDetails[self._notifHeaders['ID']]
        self._Subject.set(notifDetails[self._notifHeaders['Subject']])
        self._Due[0].set(notifDetails[self._notifHeaders['Due']].day)
        self._Due[1].set(notifDetails[self._notifHeaders['Due']].month)
        self._Due[2].set(notifDetails[self._notifHeaders['Due']].year)
        self._Due[3].set(notifDetails[self._notifHeaders['Due']].hour)
        self._Due[4].set(notifDetails[self._notifHeaders['Due']].minute)

        if notifDetails[self._notifHeaders['Category']] == "":
            self._Category.set("General")
        else:
            self._Category.set(notifDetails[self._notifHeaders['Category']])

        self._Urgency.set(notifDetails[self._notifHeaders['Urgency']])

        self._messageT.delete(1.0, 'end')
        self._messageT.insert('end', notifDetails[self._notifHeaders['Message']])

    def createNotif(self):
        #function called when creating a new notification. The initValues function initialises the tkinter variables
        # for this forms input widgets
        self.initValues()
        #rest of function creates label, entry, text and radiobutton widgets for inputting the notification's details
        self.createFrame = tk.Frame(self, bg=self._bg)
        self.createFrame.place(relx=0, rely=0, relheight=1, relwidth=1)

        self._head3 = tk.Label(self.createFrame, text="Create Notification", bg=self._bg, font=self.fonts[0])
        self._head3.place(relx=0.35, rely=0.03, relheight=0.05, relwidth=0.3)

        self._subjectL = tk.Label(self.createFrame, text="Subject", font=self.fonts[1], bg=self._bg, anchor='w')
        self._subjectL.place(relx=0.3, rely=0.1, relheight=0.05, relwidth=0.15)
        self._subjectE = tk.Entry(self.createFrame, textvariable=self._Subject, font=self.fonts[2], bg=self._bg)
        self._subjectE.place(relx=0.45, rely=0.1, relheight=0.05, relwidth=0.2)
        self._subjectForm = tk.Label(self.createFrame, text="Max. 255 char", fg='gray', font=self.fonts[3],
                                     bg=self._bg, anchor='w')
        self._subjectForm.place(relx=0.66, rely=0.1, relheight=0.05, relwidth=0.14)

        self._messageL = tk.Label(self.createFrame, text="Message", font=self.fonts[1], bg=self._bg, anchor='w')
        self._messageL.place(relx=0.3, rely=0.16, relheight=0.05, relwidth=0.15)
        self._messageSB = tk.Scrollbar(self.createFrame, orient="vertical", troughcolor='white', bg='black', bd=0)
        self._messageT = tk.Text(self.createFrame, bg=self._bg, font=self.fonts[2], wrap='word',
                                        yscrollcommand=self._messageSB.set)
        self._messageSB.config(command=self._messageT.yview)
        self._messageSB.place(relx=0.65, rely=0.16, relheight=0.15)
        self._messageT.place(relx=0.45, rely=0.16, relheight=0.15, relwidth=0.2)

        self._dueL = tk.Label(self.createFrame, text="Due Date", font=self.fonts[1], bg=self._bg, anchor='w')
        self._dueL.place(relx=0.3, rely=0.32, relheight=0.05, relwidth=0.15)
        self._dueE1 = tk.Entry(self.createFrame, textvariable=self._Due[0], bg=self._bg, font=self.fonts[2],justify='center')
        self._dueE1.place(relx=0.45, rely=0.32, relheight=0.05, relwidth=0.03)
        tk.Label(self.createFrame, text="/", bg=self._bg, font=self.fonts[2], anchor='w').place(relx=0.48, rely=0.32,
                                                                                         relheight=0.05,
                                                                                         relwidth=0.01)
        self._dueE2 = tk.Entry(self.createFrame, textvariable=self._Due[1], bg=self._bg, font=self.fonts[2],justify='center')
        self._dueE2.place(relx=0.49, rely=0.32, relheight=0.05, relwidth=0.03)
        tk.Label(self.createFrame, text="/", bg=self._bg, font=self.fonts[2], anchor='w').place(relx=0.52, rely=0.32,
                                                                                         relheight=0.05,
                                                                                         relwidth=0.01)
        self._dueE3 = tk.Entry(self.createFrame, textvariable=self._Due[2], bg=self._bg, font=self.fonts[2],justify='center')
        self._dueE3.place(relx=0.53, rely=0.32, relheight=0.05, relwidth=0.03)

        self._dueE4 = tk.Entry(self.createFrame, textvariable=self._Due[3], bg=self._bg, font=self.fonts[2],
                               justify='center')
        self._dueE4.place(relx=0.57, rely=0.32, relheight=0.05, relwidth=0.03)
        tk.Label(self.createFrame, text=":", bg=self._bg, font=self.fonts[2], anchor='w').place(relx=0.6, rely=0.32,
                                                                                         relheight=0.05,
                                                                                         relwidth=0.01)
        self._dueE5 = tk.Entry(self.createFrame, textvariable=self._Due[4], bg=self._bg, font=self.fonts[2],
                               justify='center')
        self._dueE5.place(relx=0.61, rely=0.32, relheight=0.05, relwidth=0.03)
        self._dueForm = tk.Label(self.createFrame, text="DD/MM/YYYY hh:mm", bg=self._bg, fg='gray', font=self.fonts[3],
                             anchor='w')
        self._dueForm.place(relx=0.66, rely=0.32, relheight=0.05, relwidth=0.14)

        self._catOptionsE = self.getCategories()
        self._Category.set("General")
        self._CatL = tk.Label(self.createFrame, text="Category", font=self.fonts[1], bg=self._bg, anchor='w')
        self._CatL.place(relx=0.3, rely=0.38, relheight=0.05, relwidth=0.15)
        self._CatE = ttk.Combobox(self.createFrame, textvariable=self._Category, values=self._catOptionsE)
        self._CatE.place(relx=0.45, rely=0.38, relheight=0.05, relwidth=0.2)
        self._CatForm = tk.Label(self.createFrame, text="Max. 20 char", fg='gray', font=self.fonts[3], bg=self._bg,
                                         anchor='w')
        self._CatForm.place(relx=0.66, rely=0.38, relheight=0.05, relwidth=0.14)


        self._UrgencyL = tk.Label(self.createFrame, text="Urgency", font=self.fonts[1], bg=self._bg, anchor='w')
        self._UrgencyL.place(relx=0.3, rely=0.44, relheight=0.05, relwidth=0.15)
        self._UrgencyB1 = tk.Radiobutton(self.createFrame, text="!", font=self.fonts[0], bg=self._bg, fg="green",
                                         variable=self._Urgency, value='low', indicatoron=False, relief='groove',
                                         command=lambda: self.changeUrgencyLook(one='white'), selectcolor='green')
        self._UrgencyB1.place(relx=0.5, rely=0.44, relheight=0.05, relwidth=0.05)
        self._UrgencyB2 = tk.Radiobutton(self.createFrame, text="!!", font=self.fonts[0], bg=self._bg, fg="#d68222",
                                         variable=self._Urgency, value='medium', indicatoron=False, relief='groove',
                                         command=lambda: self.changeUrgencyLook(two='white'), selectcolor='#d68222')
        self._UrgencyB2.place(relx=0.55, rely=0.44, relheight=0.05, relwidth=0.05)
        self._UrgencyB3 = tk.Radiobutton(self.createFrame, text="!!!", font=self.fonts[0], bg=self._bg, fg="red",
                                         variable=self._Urgency, value='high', indicatoron=False, relief='groove',
                                         command=lambda: self.changeUrgencyLook(three='white'), selectcolor='red')
        self._UrgencyB3.place(relx=0.6, rely=0.44, relheight=0.05, relwidth=0.05)

        #button calls the back() function when clicked which destroys this new frame and refreshes the list of
        # notifications
        self._back2B = tk.Button(self.createFrame, text="Back", bg=self._bg, font=self.fonts[2],
                                 command=lambda: self.back(self.createFrame))
        self._back2B.place(relx=0.4, rely=0.9, relheight=0.05, relwidth=0.1)

        #button calls the applyCreate function to validate the inputs and save the notification
        self._applyB = tk.Button(self.createFrame, text="Finish", bg=self._bg, font=self.fonts[2],
                                 command=self.applyCreate)
        self._applyB.place(relx=0.55, rely=0.9, relheight=0.05, relwidth=0.1)

    def applyCreate(self):
        #validation of all the inputs to ensure they exist, are valid, and conform to the set length
        if not 0 < len(self._Subject.get()) <= 255:
            tk.messagebox.showerror('Invalid Input', 'Ensure the subject is less than 256 characters.')
            return
        if not 0 < len(self._Category.get()) <= 20:
            tk.messagebox.showerror('Invalid Input', 'Ensure the category is less than 21 characters.')
            return
        if self._Category.get() == "General":
            self._Category.set("")
        if not self._Urgency.get():
            tk.messagebox.showerror('Invalid Input', 'Ensure you select an urgency level. ! is normal; !!! is '
                                                     'very high.')
            return
        if not 0 < len(self._messageT.get(1.0, 'end')) <= 4294967295:
            tk.messagebox.showerror('Invalid Input',
                                    'Enter something in the message field (must not exceed 4294967295 '
                                    'characters).')
            return

        try:
            self._dueDateFull = dt.datetime(year=self._Due[2].get(), month=self._Due[1].get(),
                                            day=self._Due[0].get(), hour=self._Due[3].get(),
                                            minute=self._Due[4].get())
        except ValueError:
            tk.messagebox.showerror('Invalid Input',
                                    'Enter a valid date and time in the form DD/MM/YYYY hh:mm for the '
                                    'notification due date.')
            return



        try:
            #if no error, adds the notification to the database or updates the existing notification depending on the
            # value of the _EditNotif boolean variable
            self.conn = MySQLConnection(**self.db)
            if self.conn.is_connected() != True:
                self._controller.log_event("Database not connected contact the network admin.",
                                           self._controller.lineno())
            self.c = self.conn.cursor()
            if not self._EditNotif:
                self.c.execute("INSERT INTO notifications (Subject, Due, Message, Category, Urgency) VALUES (%("
                               "sub)s, %(due)s, %(mess)s, %(cat)s, %(urg)s)",
                               {'sub':self._Subject.get(),'due':self._dueDateFull,'mess':self._messageT.get(1.0,'end'),
                                'cat':self._Category.get(),'urg':self._Urgency.get()})
            else:
                self.c.execute("UPDATE notifications SET Subject=%(sub)s, Due=%(due)s, Message=%(mess)s, "
                               "Category=%(cat)s, Urgency=%(urg)s WHERE id=%(id)s",
                               {'sub': self._Subject.get(),'due': self._dueDateFull,'mess': self._messageT.get(1.0, 'end'),
                                'cat': self._Category.get(),'urg': self._Urgency.get(),'id': self._EditNotif})

            tk.messagebox.showinfo("Success", "Successfully registered notification.")
        except Error as e:
            self._controller.log_event(e, self._controller.lineno())
        finally:
            self.conn.commit()
            self.conn.close()
        #back() destoys the createFrame form from the display and refreshes the list of notifications
        self.back(self.createFrame)

    def changeUrgencyLook(self, one='green', two='#d68222', three='red'):
        #when a urgency radiobutton is clicked, its foreground and background colours are inverted
        self._UrgencyB1.config(fg=one)
        self._UrgencyB2.config(fg=two)
        self._UrgencyB3.config(fg=three)

    def back(self, frame=None):
        #function destroys frames passed as a list argument, sets the program to not be editing notifications,
        # query the database for new categories and update the list of notifications
        if frame:
            frame.destroy()
        self._EditNotif = False
        self._catOptions = self.getCategories()
        self.updateItems()

    def briefNotif(self):
        #function just createsthe listbox and scrollbar widgets used to list the notifications
        self.briefFrame = tk.Frame(self, bg=self._bg)
        self.briefFrame.place(relx=0.04, rely=0.02, relheight=0.92, relwidth=0.35)
        self._head2 = tk.Label(self.briefFrame, text="Notifications", bg=self._bg, font=self.fonts[0])
        self._head2.place(relx=0.1, rely=0.02, relwidth=0.8, relheight=0.05)
        self.yscrollbar = tk.Scrollbar(self.briefFrame, orient="vertical", troughcolor="white",
                                       relief='flat', highlightthickness=0, bd=0, jump=1)
        self.notifLB = tk.Listbox(self.briefFrame, bg=self._bg, font=self.fonts[4],
                                  activestyle='none',
                               relief='flat', bd=0, highlightthickness=0, yscrollcommand=self.yscrollbar.set)
        self.yscrollbar.config(command=self.notifLB.yview)
        self.yscrollbar.place(relx=0.92, rely=0.1, relheight=0.89)
        self.notifLB.place(relx=0.05, rely=0.1, relheight=0.89, relwidth=0.87)
        #binding that calls the showItems function when a new item is selected in the listbox
        self.notifLB.bind("<<ListboxSelect>>", lambda e: self.showItems())

    def fullNotif(self):
        #function that creates all the label, scrolledtext and button widgets involved in displaying the full details
        # of a notification
        self.fullFrame = tk.LabelFrame(self.fillerFrame, text="Notification Details", bg=self._bg)
        self.fullFrame.place(relx=0, rely=0, relheight=1, relwidth=1)

        self._viewSubjectL = tk.Label(self.fullFrame, text="Subject", font=self.fonts[2], bg=self._bg)
        self._viewSubjectL.place(relx=0.05, rely=0.1, relheight=0.05, relwidth=0.2)
        self._viewSubjectT = tk.scrolledtext.ScrolledText(master=self.fullFrame, wrap=tk.WORD, bg=self._bg,
                                                          state='disabled', bd=0)
        self._viewSubjectT.place(relx=0.26, rely=0.1,relheight=0.1, relwidth=0.6)

        self._viewMessageL = tk.Label(self.fullFrame, text="Message", font=self.fonts[2], bg=self._bg)
        self._viewMessageL.place(relx=0.05, rely=0.23, relheight=0.05, relwidth=0.2)
        self._viewMessageT = tk.scrolledtext.ScrolledText(master=self.fullFrame, wrap=tk.WORD, bg=self._bg,
                                                          state='disabled', bd=0)
        self._viewMessageT.place(relx=0.26, rely=0.23,relheight=0.4, relwidth=0.6)

        self._viewDueL = tk.Label(self.fullFrame, text="Due Date", font=self.fonts[2], bg=self._bg)
        self._viewDueL.place(relx=0.05, rely=0.65, relheight=0.05, relwidth=0.2)
        self._viewDueL2 = tk.Label(self.fullFrame, text="", font=self.fonts[2], bg=self._bg,
                                   anchor='w')
        self._viewDueL2.place(relx=0.26, rely=0.65, relheight=0.05, relwidth=0.3)

        self._viewCatL = tk.Label(self.fullFrame, text="Category", font=self.fonts[2], bg=self._bg)
        self._viewCatL.place(relx=0.05, rely=0.72, relheight=0.05, relwidth=0.2)
        self._viewCatL2 = tk.Label(self.fullFrame, text="", font=self.fonts[2], bg=self._bg,
                                   anchor='w')
        self._viewCatL2.place(relx=0.26, rely=0.72, relheight=0.05, relwidth=0.6)

        self._viewUrgL = tk.Label(self.fullFrame, text="Urgency", font=self.fonts[2], bg=self._bg)
        self._viewUrgL.place(relx=0.05, rely=0.79, relheight=0.05, relwidth=0.2)
        self._viewUrgL2 = tk.Label(self.fullFrame, text="", font=self.fonts[2], bg=self._bg,
                                   anchor='w')
        self._viewUrgL2.place(relx=0.26, rely=0.79, relheight=0.05, relwidth=0.3)

        #icon button calls a function to edit the selected notification
        self.IMG2 = tk.PhotoImage(file='Images/Icons/edit.gif')
        self._editB = tk.Button(self.fullFrame, image=self.IMG2, bg=self._bg, bd=0, highlightthickness=0,
                                activebackground=self._bg, command=self.editItem)
        self._editB.image = self.IMG2
        self._editB.place(relx=0.92, rely=0.08, relwidth=0.05, relheight=0.05)

        #icon button calls a function to delete the selected notification
        self.IMG1 = tk.PhotoImage(file='Images/Icons/bin.gif')
        self._delB = tk.Button(self.fullFrame, image=self.IMG1, bg=self._bg, bd=0, highlightthickness=0,
                               activebackground=self._bg, command=self.deleteItem)
        self._delB.image = self.IMG1
        self._delB.place(relx=0.92, rely=0.02, relwidth=0.05, relheight=0.05)

        #restricts access to editing and deleting notifications to those with an access level of 3 or smaller
        if self.user['Access_Level'] > 3:
            self._editB.config(state='disabled')
            self._delB.config(state='disabled')

        #icon button calls a function to email the selected notification to all employees
        self.IMG3 = tk.PhotoImage(file='Images/Icons/send.gif')
        self._sendB = tk.Button(self.fullFrame, image=self.IMG3, bg=self._bg, bd=0,highlightthickness=0,
                               activebackground=self._bg, command=self.sendItem)
        self._sendB.image = self.IMG3
        self._sendB.place(relx=0.92, rely=0.14, relwidth=0.05, relheight=0.05)

        #removes the frame from display leaving the fillerFrame in its place
        self.fullFrame.place_forget()

    def deleteItem(self):
        #function that deletes a selected notification
        self._selected = self.notifLB.curselection()[0]
        if self._selected == "":
            #validates if notification is selected
            tk.messagebox.showerror("Select Notification", "Please select a notification to edit it.")
            return
        #deletes the notification from the listbox
        self.notifLB.delete(self._selected)
        #gets the id of the selected notification to delete it from class attriutes storing all the notifications
        self._selectedID = self._notifications[self._currNotif[self._selected][1]][self._notifHeaders['ID']]
        del self._notifications[self._currNotif[self._selected][1]]
        del self._currNotif[self._selected]
        #removing the notification from the database
        try:
            self.conn = MySQLConnection(**self.db)
            if self.conn.is_connected() != True:
                self._controller.log_event("Database not connected contact the network admin.",
                                           self._controller.lineno())
            self.c = self.conn.cursor()
            self.c.execute("DELETE FROM notifications WHERE ID=%(id)s", {'id':self._selectedID})

        except Error as e:
            self._controller.log_event(e, self._controller.lineno())
        finally:
            self.conn.commit()
            self.conn.close()

            #removing the frame with notification details from the display
            self.fullFrame.place_forget()

    def editItem(self):
        #function used to edit a notification
        self._selected = self.notifLB.curselection()[0]
        #validates if a notification is selected
        if self._selected == "":
            tk.messagebox.showerror("Select Notification", "Please select a notification to edit it.")
            return
        self._selectedItem = self._notifications[self._currNotif[self._selected][1]]
        #calls function to create the form to input notification details
        self.createNotif()
        #populate the form with details of the notification currently stored
        self.editValues(self._selectedItem)

    def sendItem(self):
        self._selected = self.notifLB.curselection()[0]
        #validates if a notification is selected
        if self._selected == "":
            tk.messagebox.showerror("Select Notification", "Please select a notification to send it.")
            return

        try:
            #selects the email of all staff members still employed
            self.conn = MySQLConnection(**self.db)
            if self.conn.is_connected() != True:
                self._controller.log_event("Database not connected contact the network admin.",
                                           self._controller.lineno())
            self.c = self.conn.cursor()
            self.c.execute("SELECT DISTINCT Email FROM staff WHERE Date_Dismissed is NULL")
            self._emailList = [elem[0] for elem in self.c.fetchall()]

        except Error as e:
            self._controller.log_event(e, self._controller.lineno())
        finally:
            self.conn.commit()
            self.conn.close()

        #Sets the sender as the logged in user's first name
        self._sender = self._controller.output_user()['First_Name']
        #constructs the email body based on the notification details
        self._subject = "Notification: " + self._notifications[self._currNotif[self._selected][1]][self._notifHeaders[
            'Subject']]

        self.messageDetails = [self._notifications[self._currNotif[self._selected][1]][self._notifHeaders['Subject']],
                               self._notifications[self._currNotif[self._selected][1]][self._notifHeaders['Message']],
                               self._notifications[self._currNotif[self._selected][1]][self._notifHeaders['Category']],
                               self._notifications[self._currNotif[self._selected][1]][self._notifHeaders['Due']],
                               self._notifications[self._currNotif[self._selected][1]][self._notifHeaders['Urgency']].upper()]

        self._receiver = self._emailList[0]
        self._body = """
You have received a reminder about the following notification for {0}:

Subject: {1}
Message: {2}
Category: {3}
The deadline for this event is set to: {4}
It has a {5} urgency.

            """.format(self.company['Company_Name'], *self.messageDetails)
        #calls a function in the Main() class to send the notification to all employees
        self._controller.email_compiler([self._sender, self._receiver, self._subject, self._body, self._emailList[1:]])

        tk.messagebox.showinfo("Sent", "Notification reminder sent to all.")

    def createSideBar(self):
        #functions creates all the widgets on the pop-up sidebar

        #creates the curved rectangle background of the pop-up sidebar
        self._sbBG = "#adccff"
        self._sbCanvas = tk.Canvas(self, bg=self._bg, bd=0, highlightthickness=0, relief='ridge')
        self._sbCanvas.place(relx=0, rely=0, relheight=1, relwidth=0.25)
        self._sbCanvas.height = self._sbCanvas.winfo_reqheight()
        self._sbCanvas.width = self._sbCanvas.winfo_reqwidth()
        self._sbCanvas.bind("<Configure>", lambda e: self.resize(self._sbCanvas, e, widthScale=True))
        self.BigRect = self.curved_rect(0, 350, 0, self._sbCanvas.height, self._sbCanvas, r=50,
                                        fill=self._sbBG)

        self._head1 = tk.Label(self._sbCanvas, text="Options", bg=self._sbBG, font=self.fonts[0])
        self._head1.place(relx=0.25, rely=0.04, relwidth=0.5, relheight=0.05)

        self._createB = tk.Button(self._sbCanvas, text="Create New Notification     +", bg=self._sbBG, font=self.fonts[
            4], bd=0,highlightthickness=0, anchor='e', command=self.createNotif)
        self._createB.place(relx=0.1, rely=0.11, relwidth=0.8, relheight=0.05)

        #if access level is greater than 3, block the user from creating a notification
        if self.access > 3:
            self._createB.config(state='disabled')


        self._catL = tk.Label(self._sbCanvas, text="Category", bg=self._sbBG, font=self.fonts[2])
        self._catL.place(relx=0.2, rely=0.18, relwidth=0.2, relheight=0.05)
        self._Cat = tk.StringVar()
        self._catO = ttk.Combobox(self._sbCanvas, textvariable=self._Cat, state="readonly")
        self._catO.place(relx=0.42, rely=0.18, relwidth=0.48, relheight=0.05)
        #binding sorts the notifications when a new category is selected
        self._catO.bind("<<ComboboxSelected>>", lambda e:self.categorise())
        self._catOptions = self.getCategories()
        self._Cat.set("Select a Category")

        self._Search = tk.StringVar()
        self._searchL = tk.Label(self._sbCanvas, text="Search", bg=self._sbBG, font=self.fonts[2])
        self._searchL.place(relx=0.2, rely=0.24, relwidth=0.2, relheight=0.05)
        self._searchE = tk.Entry(self._sbCanvas, textvariable=self._Search, font=self.fonts[2], bg='#ffffff')
        self._searchE.place(relx=0.42, rely=0.24, relwidth=0.35, relheight=0.05)
        self._searchB = tk.Button(self._sbCanvas, text="Go", bg=self._sbBG, font=self.fonts[
            2], bd=0,highlightthickness=0, activebackground=self._sbBG,command=self.search)
        self._searchB.place(relx=0.78, rely=0.24, relwidth=0.12, relheight=0.05)
        #binding searches the notifications for the search term when the user presses enter whilst focus is on the
        # search entry widget
        self._searchE.bind("<Return>", lambda e:self.search())

        self._ShowUrgency = tk.IntVar()
        self._showUrgent = tk.Checkbutton(self._sbCanvas, text="Colour Code Urgencies", variable=self._ShowUrgency, bd=0,
                                          highlightthickness=0, selectcolor='#ff8989', bg=self._sbBG,
                                          indicatoron=False, font=self.fonts[2], command=self.colourCode)
        self._showUrgent.place(relx=0.1, rely=0.31, relwidth=0.75, relheight=0.05)

        self._refreshB = tk.Button(self._sbCanvas, text="Refresh", bg=self._sbBG, font=self.fonts[
            2], command=self.updateItems)
        self._refreshB.place(relx=0.4, rely=0.4, relheight=0.05, relwidth=0.2)

        self.IMG = tk.PhotoImage(file='Images/young enterprise 2.gif')
        self._Logo1 = tk.Label(self._sbCanvas, image=self.IMG, bg=self._sbBG)
        self._Logo1.image = self.IMG
        self._Logo1.place(relx=0.1, rely=0.7, relheight=0.2, relwidth=0.8)

        #removes the pop-up sidebar from display
        self._sbCanvas.place_forget()

    def categorise(self):
        #clear the notification listbox and add only the notifications whose categories were selected in the sidebar
        self.notifLB.delete(0, 'end')
        _current = self._currNotif.copy()
        self._currNotif = []
        if not self._categorised:
            for notif in _current:
                if self._Cat.get() == "General":
                    self.addToList(self._notifications[notif[1]])
                elif self._notifications[notif[1]][self._notifHeaders['Category']] == self._Cat.get():
                    self.addToList(self._notifications[notif[1]])
        else:
            for notif in self._notifications:
                if self._Cat.get() == "General":
                    self.addToList(notif)
                elif notif[self._notifHeaders['Category']] == self._Cat.get():
                    self.addToList(notif)
        self._categorised = True
        #calls colour code function to re-colour-code the notifications if that corresponding checkbox is checked
        self.colourCode()

    def search(self):
        #function clears the notification listbox and populates it with the notifications containing the search term
        self.notifLB.delete(0, 'end')
        self._searchTerm = self._Search.get().lower()
        _current = self._currNotif.copy()
        self._currNotif = []
        if self._searchTerm == "":
            for notif in self._notifications:
                self.addToList(notif)
        else:
            for notif in _current:
                if self._searchTerm in self._notifications[notif[1]][self._notifHeaders['Subject']].lower():
                    self.addToList(self._notifications[notif[1]])
        self._categorised = False
        #calls colour code function to re-colour-code the notifications if that corresponding checkbox is checked
        self.colourCode()

    def colourCode(self):
        #colour codes all the notifications in the listbox: gray for past notifications, black for low urgency,
        # ~orange for medium urgency, and red for high urgency
        _currentDate = dt.datetime.now()
        if self._ShowUrgency.get():
            _currentDate = dt.datetime.now()
            for place, (item, index) in enumerate(self._currNotif):
                _indexItemUrg = self._notifications[index][self._notifHeaders['Urgency']]
                _indexItemDue = self._notifications[index][self._notifHeaders['Due']]
                if _indexItemDue < _currentDate:
                    self.notifLB.itemconfig(place, fg='gray')
                elif _indexItemUrg == "medium":
                    self.notifLB.itemconfig(place, fg='#d68222')
                elif _indexItemUrg == "high":
                    self.notifLB.itemconfig(place, fg='red')

        else:
            #if the colour code urgencies checkbox is not checked, just colour code the past notifications gray,
            # and everything else black
            for place, (item, index) in enumerate(self._currNotif):
                _indexItemDue = self._notifications[index][self._notifHeaders['Due']]
                if _indexItemDue < _currentDate:
                    self.notifLB.itemconfig(place, fg='gray')
                else:
                    self.notifLB.itemconfig(place, fg='black')

    def showItems(self):
        #function displays the selected notification in the fullFrame widget will the fields populated by this function
        self._selected = self.notifLB.curselection()
        #validation to check if a notification was selected
        if not self._selected:
            tk.messagebox.showerror("Select Notification", "Please select a notification to view it.")
            return
        self._selected = self._selected[0]
        self._selectedItem = self._notifications[self._currNotif[self._selected][1]]

        self._viewSubjectT.config(state='normal')
        self._viewSubjectT.delete(1.0, 'end')
        self._viewSubjectT.insert(1.0, self._selectedItem[self._notifHeaders['Subject']])
        self._viewSubjectT.config(state='disabled')

        self._viewMessageT.config(state='normal')
        self._viewMessageT.delete(1.0, 'end')
        self._viewMessageT.insert(1.0, self._selectedItem[self._notifHeaders['Message']])
        self._viewMessageT.config(state='disabled')

        self._viewDueL2.config(text=str(self._selectedItem[self._notifHeaders['Due']]))
        self._viewCatL2.config(text=str(self._selectedItem[self._notifHeaders['Category']]))
        self._viewUrgL2.config(text=str(self._selectedItem[self._notifHeaders['Urgency']]))
        self.fullFrame.place(relx=0, rely=0, relheight=1, relwidth=1)

    def updateItems(self):
        #updates the notification listbox with notifications from the database sorted in ascending order of due date
        try:
            self.conn = MySQLConnection(**self.db)
            if self.conn.is_connected() != True:
                self._controller.log_event("Database not connected contact the network admin.",
                                           self._controller.lineno())
            self.c = self.conn.cursor()
            self.c.execute("SELECT * FROM notifications ORDER BY Due")
            self._notifications = self.c.fetchall()
            self._notifHeaders = {elem[1][0]:elem[0] for elem in enumerate(self.c.description)}
            #clear listbox
            self.notifLB.delete(0, 'end')
            self._currNotif = []
            #add items to listbox
            for notif in self._notifications:
                self.addToList(notif)
            #colour code notifications if the corresponding checkbox is checked
            self.colourCode()

        except Error as e:
            self._controller.log_event(e, self._controller.lineno())
        finally:
            self.conn.commit()
            self.conn.close()
            #remove fullFrame from display
            self.fullFrame.place_forget()

    def addToList(self, item):
        #function to add a notification item to the listbox. If the notification is longer than 21 characters,
        # append an ellipsis (...) to the end and insert it alonside its due date
        self._listString = str(item[self._notifHeaders['Due']]) + ":   " + item[self._notifHeaders['Subject']][0:21]
        if len(item[self._notifHeaders['Subject']]) > 21:
            self._listString += "..."
        self.notifLB.insert('end', self._listString)
        self._currNotif.append([self._listString, self._notifications.index(item)])

    def getCategories(self):
        #function gets all the different categories used for notifications already stored in the database
        try:
            self.conn = MySQLConnection(**self.db)
            if self.conn.is_connected() != True:
                self._controller.log_event("Database not connected contact the network admin.",
                                           self._controller.lineno())
            self.c = self.conn.cursor()
            self.c.execute("SELECT DISTINCT Category FROM notifications")
            self._catOptions = [elem[0] for elem in self.c.fetchall()]
            if "" in self._catOptions:
                self._catOptions.remove("")
            self._catOptions.insert(1, "General")
            #adds these categories to the category widget on the pop-up sidebar
            self._catO.config(values=self._catOptions)

        except Error as e:
            self._controller.log_event(e, self._controller.lineno())
        finally:
            self.conn.commit()
            self.conn.close()
            return self._catOptions

    def openSideBar(self, closeOnly=False):
        #if the side-bar is displayed, adjust the position of the listbox holding frame and filler frame. Also,
        # remove the sidebar from display. Otherwise, adjust the position of the two frames again, but also display
        # the sidebar
        if self._sbCanvas.winfo_ismapped():
            self._sbCanvas.place_forget()
            self.briefFrame.place(relx=0.04, rely=0.01, relheight=0.98, relwidth=0.35)
            self.fillerFrame.place(relx=0.42, rely=0.1, relheight=0.83, relwidth=0.55)
        elif not closeOnly:
            self._sbCanvas.place(relx=0, rely=0, relheight=1, relwidth=0.25)
            self.briefFrame.place(relx=0.27, rely=0.01, relheight=0.98, relwidth=0.35)
            self.fillerFrame.place(relx=0.65, rely=0.01, relheight=0.98, relwidth=0.32)
            self._sbCanvas.focus_set()

    def curved_rect(self, x1, x2, y1, y2, canvas, r=25, **kwargs):
        #creates the curved rectangle object
        points = [x1, y1,
                  x1, y1,
                  x2-r, y1,
                  x2-r, y1,
                  x2, y1,
                  x2, y1+r,
                  x2, y1+r,
                  x2, y2-r,
                  x2, y2-r,
                  x2, y2,
                  x2-r, y2,
                  x2-r, y2,
                  x1+r, y2,
                  x1+r, y2,
                  x1, y2,
                  x1, y2,
                  x1, y2,
                  x1, y1,
                  x1, y1,
                  x1, y1]
        return canvas.create_polygon(points, smooth=True, **kwargs)

    def resize(self, canvas, event, widthScale=False):
        #resizes the sidebars when the window resizes
        canvas.hscale = float(event.height) / canvas.height
        canvas.height = event.height
        if widthScale:
            canvas.wscale = float(event.width) / canvas.width
            canvas.width = event.width
            canvas.config(width=canvas.width, height=canvas.height)
        else:
            canvas.wscale = 1
            canvas.config(height=canvas.height)
        canvas.scale("all", 0, 0, canvas.wscale, canvas.hscale)

class EventPage(tk.Frame):
    def __init__(self, parent, controller, _bg):
        #creates event page, setting values for the root window, controlling class (Main), company details,
        # background colour, font styles, user details, user access level and ID and the company's private database
        # access details
        super().__init__(parent, bg=_bg)
        self.place(x=0, y=0, relheight=1, relwidth=1)
        self._bg = _bg
        self._top = parent
        self._controller = controller
        self.company = controller.get_company_metadata()
        self.db = controller.getDBdetails()
        self.fonts = controller.getAppearance('font')
        self.user = controller.output_user()
        self.access = self.user['Access_Level']
        self.userID = self.user['ID']

        #list of all the clickable event buttons currently on page
        self.eventButtons = []

        #creates the page that displays the project's general information when clicked
        self.create_event_dashboard()

        #creates the title and the navigation bar of the page which displays all the projects as tiles
        self.MainFrame = tk.Frame(self, bg=_bg)
        self.MainFrame.place(relx=0, rely=0, relheight=1, relwidth=1)
        self._head1 = tk.Label(self.MainFrame, text="Project Management", bg=_bg, font=self.fonts[0])
        self._head1.place(relx=0.4, rely=0.02, relheight=0.05, relwidth=0.2)

        #button changes board of projects to the next 14 in the database
        self._nextPage = tk.Button(self.MainFrame, text="→", font=self.fonts[0], state='disabled', command=self.updateList)
        self._nextPage.place(relx=0.84, rely=0.9, relwidth=0.03, relheight=0.05)

        #button changes board of projects to the last 14 in the database
        self._previousPage = tk.Button(self.MainFrame, text="←", font=self.fonts[0], state='disabled',
                                       command=lambda: self.updateList(forward=False))
        self._previousPage.place(relx=0.8, rely=0.9, relwidth=0.03, relheight=0.05)

        #button refreshes board of projects
        self.IMG10 = tk.PhotoImage(file='Images/Icons/reload.gif')
        self._reloadB = tk.Button(self, image=self.IMG10, bg=self._bg, bd=0,highlightthickness=0,
                               activebackground=self._bg, command=self.refreshList)
        self._reloadB.image = self.IMG10
        self._reloadB.place(relx=0.89, rely=0.9, relwidth=0.04, relheight=0.05)

        #creates the board of project button tiles
        self.ListFrame = tk.Frame(self.MainFrame, bg=self._bg)
        self.ListFrame.place(relx=0, rely=0.08, relwidth=1, relheight=0.81)
        #creates the button to create a new project
        self.IMG = tk.PhotoImage(file='Images/Events/plus.gif')
        self._newB = tk.Button(self.ListFrame, image=self.IMG, bg=self._bg, text="New", compound='top', font=self.fonts[1],
                               bd=0,highlightthickness=0, activebackground=self._bg, command=self.new_event_page)
        self._newB.image = self.IMG
        self._newB.place(relx=0.12, rely=0.02, relwidth=0.15, relheight=0.2)

        #refreshes the board of projects
        self.refreshList()
        #creates the pop-up frame that appears when hovering over projects
        self.detailsPopUp()

    def detailsPopUp(self):
        #function that creates the pop-up frame that appears at the top right of the frame when hovering over a
        # project, displaying basic project details
        self.DetailsFrame = tk.LabelFrame(self.ListFrame, text="Project Details", bg=self._bg)
        self.DetailsFrame.place(relx=0.69, rely=0, relwidth=0.22, relheight=0.26)
        self._detailsHeader = tk.Label(self.DetailsFrame, text="Event ID: \nEvent Name: \nCreation Date: "
                                                               "\nAccess Level: \nCreator: \nStart Date:", bg=self._bg,
                                       anchor='w', justify='right', font=self.fonts[2])
        self._detailsHeader.place(relx=0, rely=0.01, relheight=0.98, relwidth=0.46)
        self._detailsField = tk.Label(self.DetailsFrame, text="", bg=self._bg, anchor='e', font=self.fonts[2])
        self._detailsField.place(relx=0.47, rely=0.01, relheight=0.98, relwidth=0.53)

        #removes pop-up frame from display
        self.DetailsFrame.place_forget()

    def showDetails(self, item):
        #function configures the pop-up frame with the project's specific details when the mouse hovers over it
        detailsString = [str(item[self._eventHeaders['Event_ID']]), "\n", item[self._eventHeaders['Event_Name']], "\n", \
                        str(item[self._eventHeaders['Creation_Date']]), "\n", str(item[self._eventHeaders['Access_Level']]),
                         "\n", item[self._eventHeaders['Creator']], "\n", str(item[self._eventHeaders['Start_Date']])]
        self._detailsField.config(text="".join(detailsString))
        self.DetailsFrame.place(relx=0.69, rely=0, relwidth=0.22, relheight=0.23)

    def create_event_dashboard(self):
        #function creates the general project profile page that is altered (to the project) and displayed when
        # clicking a project
        self.eventDash = tk.Frame(self, bg=self._bg)
        self.eventDash.place(relx=0, rely=0, relheight=1, relwidth=1)

        self._head4 = tk.Label(self.eventDash, text="SELECTED EVENT NAME", bg=self._bg, font=self.fonts[0])
        self._head4.place(relx=0.38, rely=0.05, relheight=0.05, relwidth=0.24)

        #button displays Gantt Chart
        self.IMG2 = tk.PhotoImage(file='Images/Icons/gantt.gif')
        self._chartB = tk.Button(self.eventDash, image=self.IMG2, bg=self._bg, command=self.view_chart)
        self._chartB.image = self.IMG2
        self._chartB.place(relx=0.35, rely=0.15, relheight=0.3, relwidth=0.3)

        #button edits project
        self.IMG3 = tk.PhotoImage(file='Images/Icons/edit.gif')
        self._editB = tk.Button(self.eventDash, image=self.IMG3, bg=self._bg, bd=0, highlightthickness=0)
        self._editB.image = self.IMG3
        self._editB.place(relx=0.67, rely=0.17, relheight=0.05, relwidth=0.05)

        #button deletes project
        self.IMG4 = tk.PhotoImage(file='Images/Icons/bin.gif')
        self._delB2 = tk.Button(self.eventDash, image=self.IMG4, bg=self._bg, bd=0, highlightthickness=0)
        self._delB2.image = self.IMG4
        self._delB2.place(relx=0.67, rely=0.24, relheight=0.05, relwidth=0.05)

        #button downloads all project activities to an excel spreadsheet
        self.IMG8 = tk.PhotoImage(file='Images/Icons/download.gif')
        self._downloadB = tk.Button(self.eventDash, image=self.IMG8, bg=self._bg, bd=0, highlightthickness=0,
                                    activebackground=self._bg)
        self._downloadB.image = self.IMG8
        self._downloadB.place(relx=0.67, rely=0.31, relwidth=0.05, relheight=0.05)

        #checkbutton marks project as active
        self._Active = tk.IntVar()
        self._activeB = tk.Checkbutton(self.eventDash, text="Active", bg='red', fg='white', variable=self._Active,
                                       indicatoron=False, selectcolor='green', font=self.fonts[4])
        self._activeB.place(relx=0.45, rely=0.46, relheight=0.05, relwidth=0.1)

        #labels display all information about project
        detailsFieldStr = """
Event ID: 

Creation Date: 

Minimum Access Level: 

Creator: 

Start Date: 

Number of Activities: 
"""
        self._detailsHead2 = tk.Label(self.eventDash, text=detailsFieldStr, bg=self._bg, font=self.fonts[2],
                                      justify='right', anchor='e')
        self._detailsHead2.place(relx=0.35, rely=0.52, relheight=0.32, relwidth=0.15)
        self._detailsField2 = tk.Label(self.eventDash, text="", bg=self._bg, font=self.fonts[2],
                                       justify='right', anchor='e')
        self._detailsField2.place(relx=0.5, rely=0.52, relheight=0.32, relwidth=0.15)

        #button returns back to the board of projects
        self._backB3 = tk.Button(self.eventDash, text="Back", bg=self._bg, font=self.fonts[2],
                                 command=self.eventDash.place_forget)
        self._backB3.place(relx=0.46, rely=0.89, relheight=0.05, relwidth=0.08)

    def refreshList(self):
        #function called when refreshing the board of projects
        self.startRow = -14
        #both the navigation buttons are disabled until database query proves there are more than 14 events in the
        # database
        self._previousPage.config(state='disabled')
        self._nextPage.config(state='disabled')
        #updates the board
        self.updateList()

    def updateList(self, forward=True):
        #function to updates the tiles on the board
        if forward:
            #when forward navigation arrow presses
            self.startRow += 14
        else:
            #when backward navigation arrow presses
            self.startRow -= 14
        if self.startRow < 0:
            self.startRow = 0

        #destroys all times on the board
        for e in self.eventButtons:
            e[1].destroy()
        self.eventButtons = []

        try:
            #get the next/previous 14 projects
            self.conn = MySQLConnection(**self.db)
            if self.conn.is_connected() != True:
                self._controller.log_event(self._controller.lineno(),
                                           "Database not connected contact the network admin.")
            self.c = self.conn.cursor()
            self.c.execute("SELECT COUNT(*) FROM events")
            self.eventLen = self.c.fetchall()[0][0]
            self.c.execute("SELECT * FROM events ORDER BY Active DESC, Event_ID DESC LIMIT %(s)s, 14",
                           {'s':self.startRow})
            self._eventData = {elem[0]:list(elem) for elem in self.c.fetchall()}
            self._eventHeaders = {elem[1][0]:elem[0] for elem in enumerate(self.c.description)}
        except Error as e:
            self._controller.log_event(e, self._controller.lineno())
        finally:
            self.conn.close()
        #instanciate the buttons with a corresponding project and place them on the grid
        for id,(x,y) in zip(self._eventData, ((0.32, 0.02), (0.52, 0.02), (0.12, 0.27), (0.32, 0.27),
                                            (0.52,0.27),(0.72, 0.27), (0.12, 0.52),(0.32,0.52), (0.52, 0.52),
                                            (0.72, 0.52), (0.12, 0.77),(0.32,0.77),(0.52,0.77),(0.72,0.77))):
            item = self._eventData[id]
            self.eventButtons.append([
                tk.PhotoImage(file='Images/Events/'+item[self._eventHeaders['Image']]),
                tk.Button(self.ListFrame, bg=self._bg, text=item[1],
                          compound='top', font=self.fonts[1], bd=0, highlightthickness=0,
                          activebackground=self._bg, command=lambda i=item: self.load_Event(i))
            ])
            self.eventButtons[-1][1].config(image=self.eventButtons[-1][0])
            #change it so list indices are true event if columns in database are rearranged
            self.eventButtons[-1][1].bind("<Enter>", lambda e, i=item: self.showDetails(i))
            self.eventButtons[-1][1].bind("<Leave>", lambda e: self.DetailsFrame.place_forget())
            if item[self._eventHeaders['Active']]:
                self.eventButtons[-1][1].config(bg='#98FB98')
            if self.access > item[self._eventHeaders['Access_Level']]:
                self.eventButtons[-1][1].config(state='disabled')
            self.eventButtons[-1][1].image = self.eventButtons[-1][0]
            self.eventButtons[-1][1].place(relx=x, rely=y, relwidth=0.15, relheight=0.2)

        if self.startRow <= 0:
            self._previousPage.config(state="disabled")
        else:
            self._previousPage.config(state="normal")

        if self.eventLen-self.startRow-15 < 0:
            self._nextPage.config(state="disabled")
        else:
            self._nextPage.config(state="normal")

    def new_event_page(self, edit=False, id=None):
        #function that creates the page to add activities for a project
        self.activityDict = {}
        self.IDCounter = 1
        self.EventFrame = tk.Frame(self, bg=self._bg)
        self.EventFrame.place(relx=0, rely=0, relheight=1, relwidth=1)

        #creates widgets to create the timeline page
        self._head2 = tk.Label(self.EventFrame, text="Create New Timeline", bg=self._bg, font=self.fonts[0])
        self._head2.place(relx=0.4, rely=0.02, relheight=0.05, relwidth=0.2)

        self._eventColoumns = [['Activity ID', 'Activity', 'Immediately Preceding', 'Duration', 'Description',
                                'Start Time'], [3, 8, 30, 8, 8, 8]]
        #treeview displays each activity
        self.eventTree = ttk.Treeview(self.EventFrame, height=20,
                                      columns=self._eventColoumns[0], show='headings')
        for x in range(0, 6):
            header = self._eventColoumns[0][x]
            width = self._eventColoumns[1][x]
            self.eventTree.heading(header, text=header, command=lambda
                col=header: self._controller.sortItem(col, self.eventTree))
            self.eventTree.column(header, stretch='yes', width=width)
        self.eventTree.place(relx=0.3, rely=0.1, relheight=0.5, relwidth=0.65)

        self.eventScroll = ttk.Scrollbar(self.EventFrame, orient="vertical", command=self.eventTree.yview)
        self.eventScroll.place(relx=0.95, rely=0.1, relheight=0.5)

        self.eventTree.configure(yscrollcommand=self.eventScroll.set)

        self.eventTree.bind("<Double-Button-1>", lambda e: self.edit_activity())
        #button deletes activity
        self.IMG7 = tk.PhotoImage(file='Images/Icons/bin.gif')
        self._delB = tk.Button(self.EventFrame, image=self.IMG7, bg=self._bg, bd=0,highlightthickness=0,
                               activebackground=self._bg, command=self.delete_activity)
        self._delB.image = self.IMG7
        self._delB.place(relx=0.92, rely=0.62, relwidth=0.05, relheight=0.05)

        #button returns back to the previous page
        self._backB = tk.Button(self.EventFrame, text="Back", bg=self._bg, font=self.fonts[2],
                                command=self.EventFrame.destroy)
        self._backB.place(relx=0.12, rely=0.75, relheight=0.05, relwidth=0.05)
        #button continues to editing the general event details
        self._saveB = tk.Button(self.EventFrame, text="Save & Continue", bg=self._bg, font=self.fonts[2],
                                command=lambda: self.activity_network(self.activityDict, edit=edit, id=id))
        self._saveB.place(relx=0.85, rely=0.75, relheight=0.05, relwidth=0.1)

        #frame contains fields to add activity details
        self.InsertFrame = tk.LabelFrame(self.EventFrame, text="Add Activity", bg=self._bg, )
        self.InsertFrame.place(relx=0.03, rely=0.09, relheight=0.55, relwidth=0.25)
        self._activityL = tk.Label(self.InsertFrame, text="Activity", bg=self._bg, font=self.fonts[2])
        self._activityL.place(relx=0.02, rely=0.05, relheight=0.05, relwidth=0.25)
        self._activity = tk.StringVar()
        self._activityE = tk.Entry(self.InsertFrame, textvariable=self._activity, bg=self._bg, font=self.fonts[2])
        self._activityE.place(relx=0.28, rely=0.05, relheight=0.05, relwidth=0.3)
        self._activityForm = tk.Label(self.InsertFrame, text="Max 25 char.", bg=self._bg, fg='gray', font=self.fonts[
            3], anchor='w')
        self._activityForm.place(relx=0.59, rely=0.05, relheight=0.05, relwidth=0.33)

        self._precedingL = tk.Label(self.InsertFrame, text="Preceding", bg=self._bg, font=self.fonts[2])
        self._precedingL.place(relx=0.02, rely=0.2, relheight=0.05, relwidth=0.25)
        self._preceding = tk.StringVar()
        self._precedingE = tk.Entry(self.InsertFrame, textvariable=self._preceding, bg=self._bg, font=self.fonts[2])
        self._precedingE.place(relx=0.28, rely=0.2, relheight=0.05, relwidth=0.3)
        self._precedingForm = tk.Label(self.InsertFrame, text="e.g. 5,3,8", bg=self._bg, fg='gray',
                                       font=self.fonts[3], anchor='w')
        self._precedingForm.place(relx=0.59, rely=0.2, relheight=0.05, relwidth=0.33)

        self._durationL = tk.Label(self.InsertFrame, text="Duration", bg=self._bg, font=self.fonts[2])
        self._durationL.place(relx=0.02, rely=0.35, relheight=0.05, relwidth=0.25)
        self._duration = [tk.IntVar(), tk.IntVar(), tk.IntVar()]
        self._durationDay = tk.Entry(self.InsertFrame, textvariable=self._duration[0], bg=self._bg, font=self.fonts[2])
        self._durationDay.place(relx=0.3, rely=0.35, relheight=0.05, relwidth=0.13)

        self._durationHour = tk.Entry(self.InsertFrame, textvariable=self._duration[1], bg=self._bg, font=self.fonts[2])
        self._durationHour.place(relx=0.49, rely=0.35, relheight=0.05, relwidth=0.08)
        tk.Label(self.InsertFrame, text=":", bg=self._bg, font=self.fonts[2]).place(relx=0.57, rely=0.35,
                                                                                    relheight=0.05, relwidth=0.02)
        self._durationMin = tk.Entry(self.InsertFrame, textvariable=self._duration[2], bg=self._bg, font=self.fonts[2])
        self._durationMin.place(relx=0.59, rely=0.35, relheight=0.05, relwidth=0.08)

        self._durationForm = tk.Label(self.InsertFrame, text="Days hh:mm", bg=self._bg, fg='gray',
                                      font=self.fonts[3], anchor='w')
        self._durationForm.place(relx=0.67, rely=0.35, relheight=0.05, relwidth=0.33)

        self._descriptionL = tk.Label(self.InsertFrame, text="Description", bg=self._bg, font=self.fonts[2])
        self._descriptionL.place(relx=0.02, rely=0.5, relheight=0.05, relwidth=0.25)
        self._descriptionSB = tk.Scrollbar(self.InsertFrame, orient="vertical", troughcolor="white", bg="black", bd=0)
        self._descriptionT = tk.Text(self.InsertFrame, bg=self._bg, font=self.fonts[2], wrap='word',
                                     yscrollcommand=self._descriptionSB.set)
        self._descriptionSB.config(command=self._descriptionT.yview)
        self._descriptionSB.place(relx=0.68, rely=0.5, relheight=0.1)
        self._descriptionT.place(relx=0.28, rely=0.5, relheight=0.1, relwidth=0.4)

        self._startTimeL = tk.Label(self.InsertFrame, text="Start Time", bg=self._bg, font=self.fonts[2])
        self._startTimeL.place(relx=0.02, rely=0.7, relheight=0.05, relwidth=0.25)
        self._start = [tk.IntVar(), tk.IntVar(), tk.IntVar(), tk.IntVar(), tk.IntVar()]
        self._startDay = tk.Entry(self.InsertFrame, textvariable=self._start[0], bg=self._bg, font=self.fonts[2])
        self._startDay.place(relx=0.28, rely=0.7, relheight=0.05, relwidth=0.08)
        tk.Label(self.InsertFrame, text="/", bg=self._bg, font=self.fonts[2]).place(relx=0.36, rely=0.7,
                                                                                    relheight=0.05, relwidth=0.02)
        self._startMonth = tk.Entry(self.InsertFrame, textvariable=self._start[1], bg=self._bg,
                                    font=self.fonts[2])
        self._startMonth.place(relx=0.38, rely=0.7, relheight=0.05, relwidth=0.08)
        tk.Label(self.InsertFrame, text="/", bg=self._bg, font=self.fonts[2]).place(relx=0.46, rely=0.7,
                                                                                    relheight=0.05, relwidth=0.02)
        self._startYear = tk.Entry(self.InsertFrame, textvariable=self._start[2], bg=self._bg, font=self.fonts[2])
        self._startYear.place(relx=0.48, rely=0.7, relheight=0.05, relwidth=0.1)
        self._startHour = tk.Entry(self.InsertFrame, textvariable=self._start[3], bg=self._bg, font=self.fonts[2])
        self._startHour.place(relx=0.61, rely=0.7, relheight=0.05, relwidth=0.08)
        tk.Label(self.InsertFrame, text=":", bg=self._bg, font=self.fonts[2]).place(relx=0.69, rely=0.7,
                                                                                    relheight=0.05, relwidth=0.02)
        self._startMin = tk.Entry(self.InsertFrame, textvariable=self._start[4], bg=self._bg, font=self.fonts[2])
        self._startMin.place(relx=0.71, rely=0.7, relheight=0.05, relwidth=0.08)

        self._startForm = tk.Label(self.InsertFrame, text="DD/MM/YYYY hh:mm 0=asap", bg=self._bg, fg='gray',
                                   font=self.fonts[3], anchor='w', justify='left', wraplength='53')
        self._startForm.place(relx=0.79, rely=0.63, relheight=0.2, relwidth=0.21)

        #button in frame adds activity to treeview and dictionary
        self._addActivityB = tk.Button(self.InsertFrame, text="Add activity", bg=self._bg, font=self.fonts[2],
                                       command=self.add_activity)
        self._addActivityB.place(relx=0.35, rely=0.85, relheight=0.05, relwidth=0.3)

    def fill_eventPage(self):
        #function that populates the treeview of the event page with activities of existing projects
        self.IDCounter = max(self._loadedData)+1
        for acti in self._loadedData:
            hour, minute, second = self._loadedData[acti][self._loadedHeaders['Duration']].split(":")
            duration = dt.timedelta(days=int(hour)//24, hours=(int(hour)%24)//1, minutes=int(minute)+((int(hour)%24)%1)*60)

            self.eventTree.insert('', 'end', values=(str(self._loadedData[acti][self._loadedHeaders['ArcID']]),
                                                                          self._loadedData[acti][self._loadedHeaders[
                                                                              'Activity']],
                                                                          self._loadedData[acti][
                                                                              self._loadedHeaders['Preceding']],
                                                                          duration,
                                                                          self._loadedData[acti][self._loadedHeaders[
                                                                              'Description']],
                                                                          self._loadedData[acti][self._loadedHeaders[
                                                                              'Activity_Start_Time']]))
            self.activityDict[str(self._loadedData[acti][self._loadedHeaders['ArcID']])] = [self._loadedData[acti][
                                                                                            self._loadedHeaders['Activity']],
                                                                                       self._loadedData[acti][self._loadedHeaders['Preceding']].split(','),
                                                                                       duration,
                                                                                       self._loadedData[acti][self._loadedHeaders['Description']],
                                                                                       self._loadedData[acti][self._loadedHeaders['Activity_Start_Time']]]

    def edit_activity(self):
        #function called when editing an activty
        self._selectedID = self.eventTree.item(self.eventTree.selection(), 'values')[0]
        if not self._selectedID:
            #checks activty is selected
            return
        #populates the activity creation fields with the activities details
        self._activity.set(self.activityDict[self._selectedID][0])
        preceding = ",".join(self.activityDict[self._selectedID][1])
        if preceding == '0':
            self._preceding.set('')
        else:
            self._preceding.set(preceding)
        self._duration[0].set(self.activityDict[self._selectedID][2].days)
        self._duration[1].set(self.activityDict[self._selectedID][2].seconds//3600)
        self._duration[2].set((self.activityDict[self._selectedID][2].seconds//60)%60)
        self._descriptionT.delete(1.0, 'end')
        self._descriptionT.insert('end', self.activityDict[self._selectedID][3])
        if self.activityDict[self._selectedID][4]:
            self._start[0].set(self.activityDict[self._selectedID][4].day)
            self._start[1].set(self.activityDict[self._selectedID][4].month)
            self._start[2].set(self.activityDict[self._selectedID][4].year)
            self._start[3].set(self.activityDict[self._selectedID][4].hour)
            self._start[4].set(self.activityDict[self._selectedID][4].minute)
        else:
            self._start[0].set(0)
            self._start[1].set(0)
            self._start[2].set(0)
            self._start[3].set(0)
            self._start[4].set(0)

        #finds all the rows pointing to the edited activity and point them to latest ID+1
        alteredRows = []
        for x in self.activityDict:
            if self._selectedID in self.activityDict[x][1]:
                self.activityDict[x][1].remove(self._selectedID)
                self.activityDict[x][1].append(str(self.IDCounter))
                alteredRows.append(x)
        treeData = [(self.eventTree.set(k, 'Activity ID'), k) for k in self.eventTree.get_children('')]
        for x in treeData:
            if x[0] in alteredRows:
                y = self.activityDict[x[0]].copy()
                y.insert(0, x[0])
                y[2] = ",".join(y[2])
                self.eventTree.item(x[1], values=y)

        #deletes edited activity from treeview and dictionary
        self.eventTree.delete(self.eventTree.selection())
        del self.activityDict[str(self._selectedID)]

    def delete_activity(self):
        #function deletes activity added to project
        try:
            self._selectedID = self.eventTree.item(self.eventTree.selection(), 'values')[0]
            #deletes activity from treeview and dictionary
            self.eventTree.delete(self.eventTree.selection())
            del self.activityDict[str(self._selectedID)]
            #searches through all current activities and deletes all pointers to the ID of the deleted activity
            alteredRows = []
            for x in self.activityDict:
                if self._selectedID in self.activityDict[x][1]:
                    self.activityDict[x][1].remove(self._selectedID)
                    alteredRows.append(x)
            treeData = [(self.eventTree.set(k, 'Activity ID'), k) for k in self.eventTree.get_children('')]
            for x in treeData:
                if x[0] in alteredRows:
                    y = self.activityDict[x[0]].copy()
                    y.insert(0, x[0])
                    y[2] = ",".join(y[2])
                    self.eventTree.item(x[1], values=y)
        except:
            return

    def add_activity(self):
        #function validates activity inputs and adds it to the treeview and dictionary
        if not 0 < len(self._activity.get()) < 26:
            tk.messagebox.showerror("Invalid Input", "Ensure the activity name is no more than 25 characters.")
            return

        if not 0 < len(self._descriptionT.get(1.0, 'end')) <= 4294967295:
            tk.messagebox.showerror('Invalid Input',
                                    'Enter something in the description field (must not exceed 4,294,967,295 '
                                    'characters).')
            return

        try:
            self._durationFull = dt.timedelta(days=self._duration[0].get(), hours=self._duration[1].get(),
                                              minutes=self._duration[2].get())
            if not self._durationFull:
                raise ValueError
            self._startFull = None
            for x in self._start:
                if x.get() != 0:
                    self._startFull = dt.datetime(day=self._start[0].get(), month=self._start[1].get(),
                                                  year=self._start[2].get(),
                                                  hour=self._start[3].get(), minute=self._start[4].get())
                    break

        except:
            tk.messagebox.showerror("Invalid Input", "Ensure all the values for the duration and start time are "
                                                     "integers. If you have chosen to add a start time, please ensure "
                                                     "the date is valid, otherwise leave the fields as 0. Do not "
                                                     "leave the duration as 0.")
            return
        if self._preceding.get() != "":
            try:
                self._precedingFull = [int(k) for k in self._preceding.get().split(",")]
                if max(self._precedingFull) >= self.IDCounter:
                    raise ValueError
                if min(self._precedingFull) <= 0:
                    raise ValueError
            except ValueError:
                tk.messagebox.showerror("Invalid Input", "Ensure you add each IMMEDIATELY preceding activity by "
                                                         "referring to its numerical ID and separate each with a comma. ")
                return
            self._precedingFull = [str(elem) for elem in self._precedingFull]

            for x in self._precedingFull:
                if x not in self.activityDict:
                    tk.messagebox.showerror("Invalid Input",
                                            "Ensure all the preceding activity inputs are valid activity IDs.")
                    return

        else:
            self._precedingFull = ['0']

        #add input to the treeview and dictionary
        self.eventTree.insert('', 'end', values=(str(self.IDCounter), self._activity.get(),
                                                                      ','.join(self._precedingFull),
                                                                      self._durationFull,
                                                                      self._descriptionT.get(1.0, 'end'),
                                                                      self._startFull))
        self.activityDict[str(self.IDCounter)] = [self._activity.get(), self._precedingFull, self._durationFull,
                                                  self._descriptionT.get(1.0, 'end'), self._startFull]
        self.IDCounter += 1

        #clear activity input fields
        self._activity.set('')
        self._preceding.set('')
        self._duration[0].set(0)
        self._duration[1].set(0)
        self._duration[2].set(0)
        self._descriptionT.delete(1.0, 'end')
        self._start[0].set(0)
        self._start[1].set(0)
        self._start[2].set(0)
        self._start[3].set(0)
        self._start[4].set(0)

    def activity_network(self, activities, edit=False, id=None):
        if len(activities) == 0:
            tk.messagebox.showerror("No activities", "Please add an activity first.")
            return
        all_preceding = [y for x in self.activityDict for y in self.activityDict[x][1]]
        for x in all_preceding:
            if x not in self.activityDict and x != '0':
                tk.messagebox.showerror("Invalid Input", "Ensure all the preceding activity inputs are valid activity IDs.")
                return

        #initialising / clearing variables
        self.activities = activities
        self.precedence = {}
        self.Nodes = {}
        self.Nodes2 = {}
        self.proceed = {}
        self.used = {}

        #creating the network with the nodes being before the arcs
        for x in self.activities:
            act = ','.join(self.activities[x][1])
            if act in self.precedence:
                self.precedence[act].append(x)
            else:
                self.precedence[act] = [x]
                for y in self.activities[x][1]:
                    self.Nodes[y] = self.activities[x][1]

        #finding all the arcs that hang off the end of the network
        finals = []
        for x in activities:
            if x not in self.Nodes:
                finals.append(x)
        if finals:
            self.precedence[",".join(finals)] = ['-1']
            for y in finals:
                self.Nodes[y] = finals

        # creating a network where the nodes are after the arc
        for x in self.precedence.items():
            self.proceed[",".join(x[1])] = x[0].split(",")

        for x in self.precedence.values():
            for y in x:
                self.Nodes2[y] = x

        self.network = {}
        self.absNetwork = {}
        self.criticalPath = []

        #getting all the nodes with a start time into one dictionary
        self.absoluteNodes = {}
        for x in activities:
            if activities[x][4]:
                currNode = ",".join(self.proceed[",".join(self.Nodes2[x])])
                if currNode in self.absoluteNodes:
                    if self.absoluteNodes[currNode] < activities[x][4]:
                        self.absoluteNodes[currNode] = activities[x][4]
                else:
                    self.absoluteNodes[currNode] = activities[x][4]

        self.early(['0'], start=True)
        self.used = {}
        self.late(['-1'], start=self.network[",".join(self.proceed['-1'])][0])
        self.used = {}
        self.absolute = self.startTimes()
        if self.absolute == 1:
            return
        self.new_event_details(edit=edit, id=id)
        if self.absolute:
            #if an activity has an absolute start date, display all the inputs for general project start date
            self.criticalPath = ""
            self._startDateE1.config(state='disabled')
            self._startDateE2.config(state='disabled')
            self._startDateE3.config(state='disabled')
            self._startDateE4.config(state='disabled')
            self._startDateE5.config(state='disabled')
            self.manualStart = False
        else:
            #if no activity has an absolute start date, create critical path options for the user to choose
            criticalPathOptions = ["".join(str(elem[0]+1)+' - '+str(elem[1])) for elem in enumerate(self.criticalPath)]
            self.cpOption = None
            while not self.cpOption:
                self.cpOption0 = tk.messagebox.askokcancel("Choose a critical path", "Choose a critical path option "
                                                           "and enter it in the next popup: "+
                                                           "\n".join(criticalPathOptions))
                if not self.cpOption0:
                    break
                self.cpOption = tk.simpledialog.askinteger("Enter critical path", "Enter the critical path option: ",
                                                           minvalue=1, maxvalue=len(criticalPathOptions))
            if not self.cpOption:
                self.detailsFrame.destroy()
                return
            self.criticalPath = self.criticalPath[self.cpOption-1]
            self._startDateE1.config(state='normal')
            self._startDateE2.config(state='normal')
            self._startDateE3.config(state='normal')
            self._startDateE4.config(state='normal')
            self._startDateE5.config(state='normal')
            self.manualStart = True

    def early(self, arcs, start=False):
        #getting the earliest start time
        lengths = []
        if not start:
            for x in arcs:
                lengths.append(self.used[x] + self.activities[x][2])
        else:
            lengths.append(dt.timedelta(seconds=0))

        self.network[",".join(arcs)] = [max(lengths)]

        # getting ready to pull the next arcs out and find the earliest start times for their nodes
        nextArcs = self.precedence[",".join(arcs)]
        nextArcUsed = []
        #saving arcs that have already been visited and the value of their preceding nodes
        for x in nextArcs:
            self.used[x] = max(lengths)

        #going through all the arcs and applying the function on them to find their nodes' earliest start time
        for x in nextArcs:
            ready = True
            if x in nextArcUsed or x == "-1":
                continue

            for z in self.Nodes[x]:
                if z not in self.used and z not in nextArcs:
                    ready = False
                    break
                elif z in nextArcs:
                    nextArcUsed.extend(z)
            if ready:
                self.early(self.Nodes[x])

    def late(self, arcs, start=False):
        #function getting the latest finish time
        lengths = []
        if not start:
            for x in arcs:
                lengths.append([self.used[x] - self.activities[x][2], x])
        else:
            lengths.append([start, '-1'])
        nextArcs = self.proceed[",".join(arcs)]
        minLen = min([x[0] for x in lengths])
        # a keeps track of the next node when going backward
        self.network[",".join(nextArcs)].append(minLen)
        # if late == early times
        # start is not False when finding the late time for the last node
        if minLen == self.network[",".join(nextArcs)][0] and not start:
            for x in lengths:
                if x[0] == minLen:
                    # going through the different paths
                    if self.precedence[",".join(self.Nodes[x[1]])] == ['-1']:
                        self.criticalPath.append([",".join(nextArcs), x[1]])
                    else:
                        for n in self.criticalPath.copy():
                            if n[0] == ",".join(self.Nodes[x[1]]):
                                n[0] = ",".join(nextArcs)
                                n.insert(1, x[1])
        nextArcUsed = []
        for x in nextArcs:
            self.used[x] = minLen
        for x in nextArcs:
            ready = True
            if x in nextArcUsed:
                continue

            if x in self.Nodes2:
                for z in self.Nodes2[x]:
                    if z not in self.used and z not in nextArcs:
                        ready = False
                        break
                    elif z in nextArcs:
                        nextArcUsed.extend(z)
                if ready:
                    self.late(self.Nodes2[x])

    def startTimes(self):
        # function finds the start times of each activity of the project
        self.EarliestStart = []
        self.currEarliest = []
        for x in self.absoluteNodes:
            earliest = self.absoluteNodes[x] - self.network[x][0]
            self.EarliestStart.append([earliest, x, self.absoluteNodes[x]])
            if not self.currEarliest or earliest < self.currEarliest[0]:
                self.currEarliest = [earliest, x, self.absoluteNodes[x]]

        if len(self.absoluteNodes) == 0:
            return 0

        works = True
        self.invalid = []
        for x in self.EarliestStart:
            if self.currEarliest[0] + self.network[x[1]][1] > x[2]:
                works = False
                self.invalid.append([x, self.currEarliest[0] + self.network[x[1]][1] - x[2]])
        if not works:
            self.invalidStr = []
            for x in self.invalid:
                self.invalidStr.extend(["•", x[0], " by ", x[1], " hours\n"])
            tk.messagebox.showerror("Invalid start times", "The start times you entered are not compatible with each "
                                    "other. \n As a guidline either "+self.currEarliest[1]+" is too early, "
                                    "or the following nodes are too early: \n"+"".join(self.invalidStr))
            return 1

        self.earlyStart(['0'], start=self.currEarliest[0])
        self.used = {}
        self.lateStart(['-1'], start=self.absNetwork[",".join(self.proceed['-1'])][0])
        return 2

    def earlyStart(self, arcs, start=False):
        #function getting the earliest start time
        lengths = []
        if not start:
            for x in arcs:
                lengths.append(self.used[x] + self.activities[x][2])
        else:
            lengths.append(start)

        if ",".join(arcs) in self.absoluteNodes:
            absValue = self.absoluteNodes[",".join(arcs)]
        else:
            absValue = max(lengths)
        self.absNetwork[",".join(arcs)] = [absValue]
        nextArcs = self.precedence[",".join(arcs)]
        nextArcUsed = []
        for x in nextArcs:
            self.used[x] = absValue

        for x in nextArcs:
            ready = True
            if x in nextArcUsed or x == "-1":
                continue

            for z in self.Nodes[x]:
                if z not in self.used and z not in nextArcs:
                    ready = False
                    break
                elif z in nextArcs:
                    nextArcUsed.extend(z)
            if ready:
                self.earlyStart(self.Nodes[x])

    def lateStart(self, arcs, start=False):
        #function getting the latest start time
        lengths = []
        if not start:
            for x in arcs:
                lengths.append([self.used[x] - self.activities[x][2], x])
        else:
            lengths.append([start, '-1'])
        nextArcs = self.proceed[",".join(arcs)]
        minLen = min([x[0] for x in lengths])

        if ",".join(nextArcs) in self.absoluteNodes:
            absValue = self.absoluteNodes[",".join(nextArcs)]
        else:
            absValue = minLen

        # a keeps track of the next node when going backward
        self.absNetwork[",".join(nextArcs)].append(absValue)
        nextArcUsed = []
        for x in nextArcs:
            self.used[x] = absValue
        for x in nextArcs:
            ready = True
            if x in nextArcUsed:
                continue

            if x in self.Nodes2:
                for z in self.Nodes2[x]:
                    if z not in self.used and z not in nextArcs:
                        ready = False
                        break
                    elif z in nextArcs:
                        nextArcUsed.extend(z)
                if ready:
                    self.lateStart(self.Nodes2[x])

    def new_event_details(self, edit=False, id=None):
        #creates the widgets to input the general event details
        self.detailsFrame = tk.Frame(self, bg=self._bg)
        self.detailsFrame.place(relx=0, rely=0, relheight=1, relwidth=1)
        self._head3 = tk.Label(self.detailsFrame, text="Event Details", bg=self._bg, font=self.fonts[0])
        self._head3.place(relx=0.4, rely=0.05, relheight=0.05, relwidth=0.2)

        self._Name = tk.StringVar()
        self._nameL = tk.Label(self.detailsFrame, text="Event Name", bg=self._bg, font=self.fonts[1])
        self._nameL.place(relx=0.3, rely=0.13, relheight=0.05, relwidth=0.15)
        self._nameE = tk.Entry(self.detailsFrame, textvariable=self._Name, bg=self._bg, font=self.fonts[2])
        self._nameE.place(relx=0.45, rely=0.13, relheight=0.05, relwidth=0.2)
        self._nameForm = tk.Label(self.detailsFrame, text="Max 17 char.", bg=self._bg, fg='gray',
                                  font=self.fonts[3], anchor='w')
        self._nameForm.place(relx=0.66, rely=0.13, relheight=0.05, relwidth=0.14)

        self._Access = tk.IntVar()
        self._accessL = tk.Label(self.detailsFrame, text="Access Level", bg=self._bg, font=self.fonts[1])
        self._accessL.place(relx=0.3, rely=0.2, relheight=0.05, relwidth=0.15)
        self._accessE = tk.Spinbox(self.detailsFrame, textvariable=self._Access, bg=self._bg, font=self.fonts[2],
                                   from_=1, to=5, state='readonly', increment=-1)
        self._accessE.place(relx=0.45, rely=0.2, relheight=0.05, relwidth=0.1)

        self._Start = [tk.IntVar(), tk.IntVar(), tk.IntVar(), tk.IntVar(), tk.IntVar()]
        self._startDateL = tk.Label(self.detailsFrame, text="Start Date", bg=self._bg, font=self.fonts[1])
        self._startDateL.place(relx=0.3, rely=0.27, relheight=0.05, relwidth=0.15)
        self._startDateE1 = tk.Entry(self.detailsFrame, textvariable=self._Start[0], bg=self._bg, font=self.fonts[2])
        self._startDateE1.place(relx=0.45, rely=0.27, relheight=0.05, relwidth=0.02)
        tk.Label(self.detailsFrame, text="/", bg=self._bg, font=self.fonts[2]).place(relx=0.47, rely=0.27,
                                                                                     relheight=0.05, relwidth=0.01)
        self._startDateE2 = tk.Entry(self.detailsFrame, textvariable=self._Start[1], bg=self._bg, font=self.fonts[2])
        self._startDateE2.place(relx=0.48, rely=0.27, relheight=0.05, relwidth=0.02)
        tk.Label(self.detailsFrame, text="/", bg=self._bg, font=self.fonts[2]).place(relx=0.5, rely=0.27,
                                                                                     relheight=0.05, relwidth=0.01)
        self._startDateE3 = tk.Entry(self.detailsFrame, textvariable=self._Start[2], bg=self._bg, font=self.fonts[2])
        self._startDateE3.place(relx=0.51, rely=0.27, relheight=0.05, relwidth=0.03)

        self._startDateE4 = tk.Entry(self.detailsFrame, textvariable=self._Start[3], bg=self._bg, font=self.fonts[2])
        self._startDateE4.place(relx=0.57, rely=0.27, relheight=0.05, relwidth=0.02)
        tk.Label(self.detailsFrame, text=":", bg=self._bg, font=self.fonts[2]).place(relx=0.59, rely=0.27,
                                                                                     relheight=0.05, relwidth=0.01)
        self._startDateE5 = tk.Entry(self.detailsFrame, textvariable=self._Start[4], bg=self._bg, font=self.fonts[2])
        self._startDateE5.place(relx=0.6, rely=0.27, relheight=0.05, relwidth=0.02)

        self._startDateForm = tk.Label(self.detailsFrame, text="DD/MM/YYYY HH:MM", bg=self._bg, fg='gray',
                                       font=self.fonts[3], anchor='w')
        self._startDateForm.place(relx=0.66, rely=0.27, relheight=0.05, relwidth=0.14)

        #selects an icon image
        self._Image = tk.StringVar()
        self._imageL = tk.Label(self.detailsFrame, text="Icon Image", bg=self._bg, font=self.fonts[1])
        self._imageL.place(relx=0.3, rely=0.34, relheight=0.05, relwidth=0.15)
        self._iconOpt = [''] + self.getEventIcons()
        self._imageO = ttk.OptionMenu(self.detailsFrame, self._Image, *self._iconOpt)
        self._imageO.place(relx=0.45, rely=0.34, relheight=0.05, relwidth=0.2)
        self._Image.set(self._iconOpt[1])
        self._Image.trace('w', lambda n, i, m: self.displayIcon())

        self._Icon = tk.Label(self.detailsFrame, bg=self._bg)
        #displays the selected icon
        self.displayIcon()

        #returns to the previous page
        self._backB2 = tk.Button(self.detailsFrame, text="Back", bg=self._bg, font=self.fonts[2], command=self.detailsFrame.destroy)
        self._backB2.place(relx=0.4, rely=0.7, relheight=0.05, relwidth=0.04)
        #saves event details
        self._submitB2 = tk.Button(self.detailsFrame, text="Submit", bg=self._bg, font=self.fonts[2],
                                   command=lambda: self.save_event(create=not edit, eventID=id))
        self._submitB2.place(relx=0.5, rely=0.7, relheight=0.05, relwidth=0.04)

        if edit:
            #if editing an event populate the fields
            self.fill_eventDet(id=id)

    def fill_eventDet(self, id):
        #populates the fields to set general event details
        self._Name.set(self._eventData[id][self._eventHeaders['Event_Name']])
        self._Access.set(self._eventData[id][self._eventHeaders['Access_Level']])
        self._Access.set(self._eventData[id][self._eventHeaders['Access_Level']])
        self._Image.set(self._eventData[id][self._eventHeaders['Image']])
        date = self._eventData[id][self._eventHeaders['Start_Date']]
        if date:
            self._Start[0].set(date.day)
            self._Start[1].set(date.month)
            self._Start[2].set(date.year)
            self._Start[3].set(date.hour)
            self._Start[4].set(date.minute)

    def getEventIcons(self):
        #get all icon images
        return [f for f in listdir('Images/Events/') if f.endswith('.gif')]

    def displayIcon(self):
        #function to display selected icon image if a new one is selected in the event details page
        self.IMG5 = tk.PhotoImage(file='Images/Events/' + self._Image.get())
        self._Icon.image = self.IMG5
        self._Icon.config(image=self.IMG5)
        self._Icon.place(relx=0.45, rely=0.4, relheight=0.2, relwidth=0.15)

    def save_event(self, create=False, eventID=None):
        #validate and save event to database
        if not 0 < len(self._Name.get()) <= 17:
            tk.messagebox.showerror("Invalid Input", "Ensure the event name is no more than 17 characters.")
            return
        if self.manualStart:
            try:
                self.startDateFull = dt.datetime(day=self._Start[0].get(), month=self._Start[1].get(), year=self._Start[
                    2].get(), hour=self._Start[3].get(), minute=self._Start[4].get())
            except ValueError:
                tk.messagebox.showerror('Invalid Input',
                                        'Enter a valid date and time in the form DD/MM/YYYY HH:MM for the '
                                        'start date.')
                return
            for acti in self.network:
                self.absNetwork[acti] = [self.startDateFull+self.network[acti][0],
                                         self.startDateFull+self.network[acti][1]]
        else:
            self.startDateFull = None

        try:
            #save event to database
            self.conn = MySQLConnection(**self.db)
            if self.conn.is_connected() != True:
                self._controller.log_event(self._controller.lineno(),
                                           "Database not connected contact the network admin.")
            self.c = self.conn.cursor()
            if not eventID:
                eventDB = self._controller.getID()
                eventDB += "_Event"
            else:
                self.c.execute("SELECT Event_DB FROM events WHERE Event_ID=%(id)s", {'id':eventID})
                eventDB = self.c.fetchall()[0][0]
            if create:
                self.c.execute("INSERT INTO events VALUES (NULL, %(name)s, CURRENT_TIMESTAMP, %(access)s, "
                               "%(creator)s, %(active)s, %(start)s, %(img)s, %(uniqueDB)s, %(manual)s, %(path)s)",
                               {'name':self._Name.get(), 'access':self._Access.get(), 'creator':self.userID,
                                'active':self._Active.get(), 'start':self.startDateFull, 'img':self._Image.get(),
                                'uniqueDB':eventDB, 'manual':self.manualStart, 'path':",".join(self.criticalPath)})
                self.c.execute("CREATE TABLE "+eventDB+" (ArcID int(10) NOT NULL PRIMARY KEY, Activity varchar(25) NOT "
                                "NULL, Preceding varchar(50) NOT NULL, Proceeding varchar(50) NOT NULL, Duration varchar("
                                "25) NOT NULL, Activity_Start_Time datetime, Description longtext NOT NULL, "
                               "Start_Time datetime NOT NULL, Finish_Time datetime NOT NULL, Latest_Finish_Time "
                               "datetime NOT NULL)")
            else:
                self.c.execute("UPDATE events SET Event_Name=%(name)s, Access_Level=%(access)s, "
                               "Active=%(active)s, Start_Date=%(start)s, Image=%(img)s, Manual_Start=%(manual)s, "
                               "Critical_Path=%(path)s WHERE Event_ID=%(id)s",
                               {'name':self._Name.get(), 'access':self._Access.get(),
                                'active':self._Active.get(), 'start':self.startDateFull, 'img':self._Image.get(),
                                'manual':self.manualStart, 'path':",".join(self.criticalPath), 'id':eventID})
                self.c.execute("TRUNCATE TABLE "+eventDB)

            for acti in self.activities:
                beforeNode = ",".join(self.activities[acti][1])
                self.c.execute("INSERT INTO "+eventDB+" VALUES (%(id)s, %(name)s, %(before)s, %(after)s, "
                               "%(duration)s, %(actiStart)s, %(descrip)s, %(start)s, %(finish)s, %(lateFinish)s)",
                               {'id':acti, 'name':self.activities[acti][0], 'before':beforeNode,
                                'after':",".join(self.Nodes[acti]), 'duration':self.activities[acti][2],
                                'actiStart':self.activities[acti][4], 'descrip':self.activities[acti][3],
                                'start':self.absNetwork[beforeNode][0],
                                'finish':self.absNetwork[beforeNode][0]+self.activities[acti][2],
                                'lateFinish':self.absNetwork[",".join(self.Nodes[acti])][1]})

            tk.messagebox.showinfo("Success", "Event saved successfully.")


        except Error as e:
            self._controller.log_event(e, self._controller.lineno())
        finally:
            self.conn.commit()
            self.conn.close()
            #destroy, from display, all the frames to create a project
            self.detailsFrame.destroy()
            self.EventFrame.destroy()
            #refresh project board
            self.refreshList()
            if not create:
                #if editing a project, navigate back to management dashboard
                self.eventDash.place_forget()

    def editEvent(self, id):
        #function to edit existing project
        self.new_event_page(edit=True, id=id)
        self.fill_eventPage()

    def delEvent(self, id, tbName):
        #function to delete existing project from database
        if not tk.messagebox.askyesno("Sure", "Are you sure you want to delete this event?"):
            return
        try:
            self.conn = MySQLConnection(**self.db)
            if self.conn.is_connected() != True:
                self._controller.log_event(self._controller.lineno(),
                                           "Database not connected contact the network admin.")
            self.c = self.conn.cursor()
            self.c.execute("DELETE FROM events WHERE Event_ID=%(id)s", {'id':id})
            self.c.execute("DROP TABLE "+tbName)
        except Error as e:
            self._controller.log_event(e, self._controller.lineno())
        finally:
            self.conn.commit()
            self.conn.close()
            #refresh project board and navigate back to it
            self.refreshList()
            self.eventDash.place_forget()

    def load_Event(self, item):
        #function to load event details to configure the event dashboard frame to the selected project
        try:
            self.conn = MySQLConnection(**self.db)
            if self.conn.is_connected() != True:
                self._controller.log_event(self._controller.lineno(),
                                           "Database not connected contact the network admin.")
            self.c = self.conn.cursor()
            self.c.execute("SELECT * FROM " + item[self._eventHeaders['Event_DB']])
            self._loadedData = {elem[0]:elem for elem in self.c.fetchall()}
            self._loadedHeaders = {elem[1][0]:elem[0] for elem in enumerate(self.c.description)}

            self._head4.config(text=item[self._eventHeaders['Event_Name']])
            self._Active.set(item[self._eventHeaders['Active']])
            self._activeB.config(command=lambda: self.activateEvent(item))
            self._delB2.config(command=lambda: self.delEvent(item[self._eventHeaders['Event_ID']],
                                                            item[self._eventHeaders['Event_DB']]))
            self._downloadB.config(command=lambda: self.downloadDB(str(item[self._eventHeaders['Event_ID']])+" - "+
                                                            item[self._eventHeaders['Event_Name']]))
            self._editB.config(command=lambda: self.editEvent(item[self._eventHeaders['Event_ID']]))
            self._detailsField2.config(text="\n" + str(item[self._eventHeaders['Event_ID']]) + "\n\n" +
                                            str(item[self._eventHeaders['Creation_Date']]) + "\n\n" +
                                            str(item[self._eventHeaders['Access_Level']]) + "\n\n" +
                                            str(item[self._eventHeaders['Creator']]) + "\n\n" +
                                            str(item[self._eventHeaders['Start_Date']]) + "\n\n" +
                                            str(len(self._loadedData)))

            self.eventDash.place(relx=0, rely=0, relheight=1, relwidth=1)
            self.eventDash.tkraise()
            self._chartB.config(command=lambda:self.view_chart(item, item[self._eventHeaders['Event_Name']]))

        except Error as e:
            self._controller.log_event(e, self._controller.lineno())
        finally:
            self.conn.commit()
            self.conn.close()

    def view_chart(self, item, name):
        #function to create and display the Gantt chart in a browser
        df = []
        currentActiList = self._loadedData.copy()
        criticalPath = item[self._eventHeaders['Critical_Path']].split(",")
        alternating = 'Alternating1'
        counter = 0
        self.chart_reference()
        self.chartData = []
        self._downloadChartRef.config(
            command=lambda: self.downloadChart(str(item[self._eventHeaders['Event_ID']]) + " - " +
                                               item[self._eventHeaders['Event_Name']]))
        for acti in criticalPath[1:]:
            df.append({'Task':"Critical Path", 'Start':self._loadedData[
                int(acti)][self._loadedHeaders['Start_Time']], 'Finish':self._loadedData[int(acti)][self._loadedHeaders[
                'Finish_Time']], 'Resource':alternating})
            if alternating == 'Alternating1':
                alternating = 'Alternating2'
            else:
                alternating = 'Alternating1'
            self._refTree.insert("", "end", values=[counter, 'Yes', self._loadedData[int(acti)][self._loadedHeaders[
                'Activity']], self._loadedData[int(acti)][self._loadedHeaders['Duration']], self._loadedData[int(
                acti)][self._loadedHeaders['Start_Time']], self._loadedData[int(acti)][self._loadedHeaders[
                'Description']]])
            self.chartData.append([counter, 'Yes', self._loadedData[int(acti)][self._loadedHeaders[
                'Activity']], self._loadedData[int(acti)][self._loadedHeaders['Duration']], self._loadedData[int(
                acti)][self._loadedHeaders['Start_Time']], self._loadedData[int(acti)][self._loadedHeaders[
                'Description']]])
            counter += 1
            del currentActiList[int(acti)]
        for acti in currentActiList:
            df.append({'Task':self._loadedData[int(acti)][self._loadedHeaders['Activity']], 'Start':self._loadedData[
                int(acti)][self._loadedHeaders['Start_Time']], 'Finish':self._loadedData[int(acti)][self._loadedHeaders[
                'Latest_Finish_Time']], 'Resource':'Latest'})
            df.append({'Task':self._loadedData[int(acti)][self._loadedHeaders['Activity']], 'Start':self._loadedData[
                int(acti)][self._loadedHeaders['Start_Time']], 'Finish':self._loadedData[int(acti)][self._loadedHeaders[
                'Finish_Time']], 'Resource':'Earliest'})
            self._refTree.insert("", "end", values=[counter, 'No', self._loadedData[int(acti)][self._loadedHeaders[
                'Activity']], self._loadedData[int(acti)][self._loadedHeaders['Duration']], self._loadedData[int(
                acti)][self._loadedHeaders['Start_Time']], self._loadedData[int(acti)][self._loadedHeaders[
                'Description']]])
            self.chartData.append([counter, 'No', self._loadedData[int(acti)][self._loadedHeaders[
                'Activity']], self._loadedData[int(acti)][self._loadedHeaders['Duration']], self._loadedData[int(
                acti)][self._loadedHeaders['Start_Time']], self._loadedData[int(acti)][self._loadedHeaders[
                'Description']]])
            counter += 1

        #colour code the different bars
        colors = {'Earliest': 'rgb(255, 94, 94)',
                  'Latest': 'rgb(204, 222, 249)',
                  'Alternating1': 'rgb(242, 153, 65)',
                  'Alternating2': 'rgb(241, 221, 64)'}
        fig = create_gantt(df, colors=colors, index_col='Resource', showgrid_x=True,
                              showgrid_y=True, show_colorbar=False, group_tasks=True, title=name)
        fig['layout']['margin'] = {'l': 200}
        fig['layout']['hovermode'] = "closest"
        fig['layout']['autosize'] = True

        #plot the graph
        plot(fig)

    def downloadDB(self, sheetName):
        #function to download database of all activities to excel spreadsheet
        style0 = easyxf("align:horiz centre")
        style1 = easyxf('align:horiz center', num_format_str='DD-MMM-YY')
        style2 = easyxf('pattern: pattern solid, fore_color yellow; font: bold on;align:horiz center')

        self.book = Workbook()
        self.sheet = self.book.add_sheet(sheetName)

        for col, CI in list(self._loadedHeaders.items()):
            self.sheet.write(0, CI, col, style2)
        for r, row in enumerate(list(self._loadedData.values())):
            for c, col in enumerate(row):
                if isinstance(col, dt.date):
                    self.sheet.write(r + 1, c, col, style1)
                else:
                    self.sheet.write(r + 1, c, col, style0)

        self.file = filedialog.asksaveasfilename(title="Save Event Data", initialdir='Log/',
                                                 defaultextension=".xls")
        if self.file:
            self.book.save(self.file)

    def downloadChart(self, sheetname):
        #function to download the chart reference table to an excel spreadsheet
        style0 = easyxf("align:horiz centre")
        style1 = easyxf('align:horiz center', num_format_str='DD-MMM-YY')
        style2 = easyxf('pattern: pattern solid, fore_color yellow; font: bold on;align:horiz center')

        self.book = Workbook()
        self.sheet = self.book.add_sheet(sheetname+" Ref.")

        for CI, col in enumerate(('Order', 'Critical', 'Activity', 'Duration (hh:mm:ss)', 'Start Time', 'Description')):
            self.sheet.write(0, CI, col, style2)
        for r, row in enumerate(self.chartData):
            for c, col in enumerate(row):
                if isinstance(col, dt.date):
                    self.sheet.write(r + 1, c, col, style1)
                else:
                    self.sheet.write(r + 1, c, col, style0)

        self.file = filedialog.asksaveasfilename(title="Save Event Data", initialdir='Log/',
                                                 defaultextension=".xls")
        if self.file:
            self.book.save(self.file)

    def chart_reference(self):
        #function creates treeview when viewing Gantt chart to show what each critical path activity is
        self.referenceFrame = tk.LabelFrame(self, text="Chart Reference", bg=self._bg)
        self.referenceFrame.place(relx=0.03, rely=0.03, relheight=0.94, relwidth=0.94)
        self._refHeaders = ('Order', 'Critical', 'Activity', 'Duration (hh:mm:ss)', 'Start Time', 'Description')
        self._refTree = ttk.Treeview(self.referenceFrame, height=25, columns=self._refHeaders, show='headings')
        for header in self._refHeaders:
            self._refTree.heading(header, text=header, command=lambda
                col=header: self._controller.sortItem(col, self._refTree))
            self._refTree.column(header, stretch='yes', width=8)
        self._refTree.place(relx=0.01, rely=0.05, relheight=0.93, relwidth=0.92)

        self._closeB = tk.Button(self.referenceFrame, text="X", bg=self._bg, font=self.fonts[5],
                                 command=self.referenceFrame.destroy,bd=0, highlightthickness=0)
        self._closeB.place(relx=0.94, rely=0.01, relheight=0.05, relwidth=0.05)

        #button to download table
        self.IMG9 = tk.PhotoImage(file='Images/Icons/download.gif')
        self._downloadChartRef = tk.Button(self.referenceFrame, image=self.IMG9, bg=self._bg, bd=0,highlightthickness=0,
                               activebackground=self._bg)
        self._downloadChartRef.image = self.IMG9
        self._downloadChartRef.place(relx=0.94, rely=0.08, relwidth=0.05, relheight=0.05)

    def activateEvent(self, item):
        #function activates project
        try:
            self.conn = MySQLConnection(**self.db)
            if self.conn.is_connected() != True:
                self._controller.log_event(self._controller.lineno(),
                                           "Database not connected contact the network admin.")
            self.c = self.conn.cursor()
            self.c.execute("UPDATE events SET Active=%(active)s WHERE Event_ID=%(id)s", {'id':item[
                self._eventHeaders['Event_ID']], 'active':self._Active.get()})
        except Error as e:
            self._controller.log_event(e, self._controller.lineno())
        finally:
            self.conn.commit()
            self.conn.close()
            self.refreshList()

class StocksPage(tk.Frame):
    def __init__(self, parent, controller, _bg):
        # creates stocks page, setting values for the root window, controlling class (Main), company details,
        # background colour, font styles, user details, user access level and ID and the general and company's private
        # database access details
        super().__init__(parent, bg=_bg)
        self.place(x=0, y=0, relheight=1, relwidth=1)
        self._bg = _bg
        self._top = parent
        self._controller = controller
        self.company = self._controller.get_company_metadata()
        self.user = controller.output_user()
        self.db = self._controller.getDBdetails()
        self.ini_db = self._controller.getInitDB()
        self.fonts = controller.getAppearance('font')

        #creates widgets to manage and display stocks/share data
        self._head1 = tk.Label(self, text="Company Stocks", bg=_bg, font=self.fonts[0])
        self._head1.place(relx=0.4, rely=0.01, relheight=0.05, relwidth=0.2)

        self._stockHeaders = ("Share ID", "Shareholder", "Quantity", "Share Price", "Date", "Percentage")
        #treeview is like a table displaying the distribution of shares
        self.stockTree = ttk.Treeview(self, height=20, columns=self._stockHeaders, show='headings')
        for heading in self._stockHeaders:
            self.stockTree.heading(heading, text=heading, command=lambda
                col=heading: self._controller.sortItem(col, self.stockTree))
            self.stockTree.column(heading, stretch='yes', width=8)

        self.stockTree.place(relx=0.05, rely=0.15, relheight=0.7, relwidth=0.65)

        self.treeScroll = ttk.Scrollbar(self, orient="vertical", command=self.stockTree.yview)
        self.treeScroll.place(relx=0.7, rely=0.15, relheight=0.7)

        self.stockTree.configure(yscrollcommand=self.treeScroll.set)
        #search element searches the treeview
        self.Search_ = [tk.StringVar(), tk.StringVar()]
        self._searchL = tk.Label(self, text="Search", font=self.fonts[1], bg=_bg)
        self._searchL.place(relx=0.05, rely=0.91, relwidth=0.08, relheight=0.05)
        self._searchO = ttk.OptionMenu(self, self.Search_[0], '',*self._stockHeaders)
        self._searchO.place(relx=0.13, rely=0.91, relwidth=0.09, relheight=0.05)
        self.Search_[0].set(self._stockHeaders[0])
        self._searchE = tk.Entry(self, textvariable=self.Search_[1], bg=_bg, font=self.fonts[2])
        self._searchE.place(relx=0.22, rely=0.91, relwidth=0.13, relheight=0.05)
        self._searchB = tk.Button(self, text="Go", font=self.fonts[2], bg=_bg, command=lambda:
        self._controller.searchItem(
            self.Search_[0].get(), self.Search_[1].get(), self.stockTree, self.treeChild), activebackground=_bg,
                                  highlightthickness=0, bd=0)
        self._searchB.place(relx=0.35, rely=0.91, relheight=0.05, relwidth=0.05)

        self._searchE.bind("<Return>", lambda e: self._controller.searchItem(
            self.Search_[0].get(), self.Search_[1].get(), self.stockTree, self.treeChild))

        #displays general stats about company's shares
        self._head3 = tk.Label(self, text="Company Stock Details", bg=_bg, font=self.fonts[2])
        self._head3.place(relx=0.768, rely=0.2, relheight=0.05, relwidth=0.142)
        self._statsHeader = tk.Label(self, text="Total Shares: \n\nShare Price: \n\nMarket Cap: ", font=self.fonts[
            2], bg=_bg, anchor='e')
        self._statsHeader.place(relx=0.77, rely=0.27, relheight=0.17, relwidth=0.1)
        self._statsDetails = tk.Label(self, text="", font=self.fonts[ 2], bg=_bg, anchor='w')
        self._statsDetails.place(relx=0.87, rely=0.27, relheight=0.17, relwidth=0.1)

        #updates company's stock details
        self.updateStockDetails()

        #buy new share
        self.IMG2 = tk.PhotoImage(file='Images/Icons/add.gif')
        self._addB = tk.Button(self, image=self.IMG2, bg=self._bg, bd=0,highlightthickness=0,
                               activebackground=self._bg, command=self.addStock)
        self._addB.image = self.IMG2
        self._addB.place(relx=0.80, rely=0.47, relwidth=0.07, relheight=0.07)

        #buy share with paypal
        self.IMG4 = tk.PhotoImage(file='Images/Icons/paypal.gif')
        self._paypalB = tk.Button(self, image=self.IMG4, bg=self._bg, bd=0,highlightthickness=0,
                               activebackground=self._bg, command=lambda: self.addStock(paypal=True))
        self._paypalB.image = self.IMG4
        self._paypalB.place(relx=0.86, rely=0.46, relwidth=0.07, relheight=0.1)

        #delete share
        self.IMG1 = tk.PhotoImage(file='Images/Icons/bin.gif')
        self._delB = tk.Button(self, image=self.IMG1, bg=self._bg, bd=0, highlightthickness=0,
                               activebackground=self._bg, command=self.delStock)
        self._delB.image = self.IMG1
        self._delB.place(relx=0.62, rely=0.91, relwidth=0.05, relheight=0.05)

        #edit share details
        self.IMG3 = tk.PhotoImage(file='Images/Icons/edit.gif')
        self._editB = tk.Button(self, image=self.IMG3, bg=self._bg, bd=0,highlightthickness=0,
                               activebackground=self._bg, command=self.editStockDet)
        self._editB.image = self.IMG3
        self._editB.place(relx=0.95, rely=0.2, relwidth=0.05, relheight=0.05)

        #download share table to excel spreadsheet
        self.IMG5 = tk.PhotoImage(file='Images/Icons/download.gif')
        self._downloadB = tk.Button(self, image=self.IMG5, bg=self._bg, bd=0,highlightthickness=0,
                               activebackground=self._bg, command=self.downloadDB)
        self._downloadB.image = self.IMG5
        self._downloadB.place(relx=0.90, rely=0.2, relwidth=0.05, relheight=0.05)

        if self.user['Access_Level'] > 3:
            #disable the delete, edit and download button is access level is greater than 3
            self._delB.config(state='disabled')
            self._editB.config(state='disabled')
            self._downloadB.config(state='disabled')

        #refresh share table
        self.IMG6 = tk.PhotoImage(file='Images/Icons/reload.gif')
        self._reloadB = tk.Button(self, image=self.IMG6, bg=self._bg, bd=0,highlightthickness=0,
                               activebackground=self._bg, command=self.updateStocks)
        self._reloadB.image = self.IMG6
        self._reloadB.place(relx=0.73, rely=0.15, relwidth=0.04, relheight=0.05)

        #button to display share distribution as pie chart
        self.viewPie = [tk.IntVar(), tk.IntVar(), tk.IntVar()]
        self.IMG4 = tk.PhotoImage(file='Images/Icons/pie.gif')
        self._pieB = tk.Checkbutton(self, image=self.IMG4, variable=self.viewPie[0], bd=0,
                                          highlightthickness=0, selectcolor='#ff8989', bg=self._bg,
                                    indicatoron=False, command=lambda: self.createPie(hideLabels=self.viewPie[1].get(),
                                                                                      window=self.viewPie[2].get()))
        self._pieB.image = self.IMG4
        self._pieB.place(relx=0.77, rely=0.64, relwidth=0.17, relheight=0.23)

        #checkbutton to display pie chart without chart labels
        self._hideLabelB = tk.Checkbutton(self, text="Hide Labels", variable=self.viewPie[1], bg=self._bg)
        self._hideLabelB.place(relx=0.8, rely=0.88, relheight=0.05, relwidth=0.1)

        # checkbutton to display pie chart in new window
        self._windowB = tk.Checkbutton(self, text="New window", variable=self.viewPie[2], bg=self._bg)
        self._windowB.place(relx=0.8, rely=0.93, relheight=0.05, relwidth=0.1)

        if not self._stockDetData:
            #if no details about stock currently registered, display page to add those details
            self.setStockDetails()
        else:
            #if stock details, set the label to display them
            self._statsDetails.config(text=str(self._stockDetData[self._stockDetHeaders['Total_Shares']])+"\n\n"+str(
                self._stockDetData[self._stockDetHeaders['Share_Price']])+"\n\n"+str(self._stockDetData[
                self._stockDetHeaders['Market_Capitalisation']]))
            self.updateStocks()

    def addStock(self, paypal=False):
        #function used to add a new stock
        sharesPurchasable = int(self._stockDetData[self._stockDetHeaders['Total_Shares']]) - self.boughtStocks
        quantity = tk.simpledialog.askinteger("Quantity", "Enter the number of shares you want to buy: ", minvalue=1,
                                             maxvalue=int(sharesPurchasable))
        if not quantity:
            return
        name = tk.simpledialog.askstring("Your Name", "Enter your name as the shareholder: ")
        if not name:
            return

        if paypal:
            #if buying with paypal, call the PayPal_Pay function to buy stock with PayPal
            self.transaction = self.PayPal_Pay(quantity, self._stockDetData[self._stockDetHeaders['Share_Price']], name)
            if not self.transaction:
                tk.messagebox.showerror("Payment Failed", "Payment did not go through.")
                return
        else:
            self.transaction = [None, None]

        try:
            #adds bought stocks to database
            self.conn = MySQLConnection(**self.db)
            if self.conn.is_connected() != True:
                self._controller.log_event(self._controller.lineno(),
                                           "Database not connected contact the network admin.")
            self.c = self.conn.cursor()
            self.c.execute("INSERT INTO stock VALUES (NULL, %(name)s, %(quantity)s, %(price)s, CURRENT_TIMESTAMP, "
                           "%(paypal)s, %(pptax)s)",
                           {'name':name, 'quantity':quantity, 'price':self._stockDetData[self._stockDetHeaders[
                               'Share_Price']], 'paypal':self.transaction[0], 'pptax':self.transaction[1]})
            tk.messagebox.showinfo("Success", "Shares successfully bought!")
        except Error as e:
            self._controller.log_event(e, self._controller.lineno())
        finally:
            self.conn.commit()
            self.conn.close()
            #refreshes stock treeview
            self.updateStocks()

    def PayPal_Pay(self, quantity, price, name):
        #function carries out transaction with PayPal
        total = quantity * price
        if tk.messagebox.askyesno("Add fee", "Do you want to add the 3.4% + 20p PayPal fee to the purchase price?"):
            #adds fee on top of stock purchase cost
            onlineTax = total * 0.034
            onlineTax += 0.2
            sale_price = round(total + onlineTax, 2)
            onlineTax = round(onlineTax, 2)
            paypalTax = True
        else:
            onlineTax = 0
            sale_price = round(total, 2)
            paypalTax = False

        #creates payment
        payment = paypal.Payment({
            "intent": "sale",
            "redirect_urls": {
                "return_url": website,
                "cancel_url": website
            },
            "payer": {
                "payment_method": "paypal",
            },
            "transactions": [
                {
                    "amount": {
                        "total": str(sale_price),
                        "currency": "GBP",
                        "details": {
                            "subtotal": str(total),
                            "handling_fee": str(onlineTax),
                        }
                    },
                    "description": "Buy a stock; invest in us!",
                    "soft_descriptor": self.company['Company_Name']+" - shares",
                    "item_list": {
                        "items": [
                            {
                                "name": self.company['Company_Name']+" Shares for "+name,
                                "quantity": str(quantity),
                                "price": str(price),
                                "sku": "Shares",
                                "currency": "GBP"
                            }
                        ]
                    }
                }
            ]
        })

        if payment.create():
            self._controller.log_event("Payment[%s] created successfully" % (payment.id), self._controller.lineno())
            # Redirect the user to given approval url
            for link in payment.links:
                if link.rel == "approval_url":
                    # Convert to str to avoid google appengine unicode issue
                    # https://github.com/paypal/rest-api-sdk-python/pull/58
                    approval_url = str(link.href)
                    self._controller.log_event("Redirect for approval: %s" % (approval_url), self._controller.lineno())
                    webb.open(approval_url)
        else:
            self._controller.log_event("Error while creating payment:", self._controller.lineno())
            self._controller.log_event(payment.error, self._controller.lineno())
            return

        #window pauses the program until the user enters their details into the payment gateway page
        if tk.messagebox.askyesno("Authorise payment", "Your browser should have opened a paypal payment page. Have "
                                                       "you entered your details and authorised the payment?"):
            # Payment ID obtained when creating the payment (following redirect)
            payment = paypal.Payment.find(payment.id)
            payerid = payment.payer.payer_info.payer_id

            # Execute payment with the payer ID from the create payment call (following redirect)
            if payment.execute({"payer_id": payerid}):
                self._controller.log_event("Payment[%s] execute successfully" % (payment.id))
                transactionid = payment.transactions[0].related_resources[0].sale.id
                return [transactionid, paypalTax]
            else:
                self._controller.log_event(payment.error)
        else:
            return

    def PayPal_Refund(self, id, quantity, price, tax):
        #called when deleting a stock to refund paypal payment made
        if tax:
            #if fee added to payment, refund fee
            if quantity == int(self.selectionData[2]):
                refundAmount = quantity*price*1.034
                refundAmount += 0.2
            else:
                refundAmount = quantity*price*1.034
        else:
            refundAmount = quantity*price
        refundAmount = round(refundAmount, 2)

        sale = paypal.Sale.find(id)
        # Make Refund API call
        refund = sale.refund({
            "amount": {
                "total": str(refundAmount),
                "currency": "GBP"
            }
        })
        # Check refund status
        if refund.success():
            self._controller.log_event("Refund[%s] Success" % (refund.id))
            return True
        else:
            self._controller.log_event("Unable to Refund")
            self._controller.log_event(refund.error)
            return

    def delStock(self):
        #function used to delete stocks
        if not self.stockTree.selection():
            #checks if item selected in treeview
            tk.messagebox.showerror("Select stock", "Please select a stock you wish to delete.")
            return
        self.selectionData = self.stockTree.item(self.stockTree.selection(), 'values')
        quantity = tk.simpledialog.askinteger("Quantity", "Enter the number of shares you want to delete: ", minvalue=1,
                                             maxvalue=int(self.selectionData[2]))
        if not quantity:
            return
        try:
            #gets the specific stock detail from database
            self.conn = MySQLConnection(**self.db)
            if self.conn.is_connected() != True:
                self._controller.log_event(self._controller.lineno(),
                                           "Database not connected contact the network admin.")
            self.c = self.conn.cursor()
            self.c.execute("SELECT PayPal, Price_Per_Share, PayPal_Tax FROM stock WHERE Share_ID=%(id)s",
                           {'id':self.selectionData[0]})
            transaction = self.c.fetchall()[0]
            if transaction[0]:
                #if there is a transaction id from paypal, ask to refund paypal purhcase
                if tk.messagebox.askyesno("Refund Paypal", "Do you want to refund the amount paid through paypal?"):
                    #if yes, call the PayPal_Refund function
                    if not self.PayPal_Refund(transaction[0], quantity, transaction[1], transaction[2]):
                        tk.messagebox.showerror("Refund unavailable", "The refund did not process. Please contact the "
                                                                 "system administator for further support.")
                        return
            #delete the stock from the database
            if quantity == int(self.selectionData[2]):
                self.c.execute("DELETE FROM stock WHERE Share_ID=%(id)s", {'id':self.selectionData[0]})
            else:
                self.c.execute("UPDATE stock SET Quantity=%(quantity)s WHERE Share_ID=%(id)s",
                               {'quantity': int(self.selectionData[2])-quantity, 'id': self.selectionData[0]})

            tk.messagebox.showinfo("Success", "Shares deleted successfully!")
        except Error as e:
            self._controller.log_event(e, self._controller.lineno())
        finally:
            self.conn.commit()
            self.conn.close()
            #refresh stock treeview
            self.updateStocks()

    def updateStocks(self):
        #function that updates the stock treeview

        #clear the stock treeview
        self.stockTree.delete(*self.stockTree.get_children())
        try:
            #get all stock holdings from database and add it to the treeview
            self.conn = MySQLConnection(**self.db)
            if self.conn.is_connected() != True:
                self._controller.log_event(self._controller.lineno(),
                                           "Database not connected contact the network admin.")
            self.c = self.conn.cursor()
            self.c.execute("SELECT * FROM stock")

            self._stockHeadersList = [elem[0] for elem in self.c.description]
            self._stockHeaders = {elem[1]:elem[0] for elem in enumerate(self._stockHeadersList)}
            self._stockData = {elem[0]:list(elem) for elem in self.c.fetchall()}
            self._stockDataList = list(self._stockData.values())
            self._stockHeaders['Percent'] = len(self._stockHeaders)
            self._stockHeadersList.append('Percent')

            outstanding = self._stockDetData[self._stockDetHeaders['Total_Shares']]
            for index, item in enumerate(self._stockData):
                percentage = self._stockData[item][self._stockHeaders['Quantity']] / outstanding
                percentage = round(percentage*100, 1)
                self._stockData[item].append(percentage)
                self.stockTree.insert("", "end", values=[self._stockData[item][self._stockHeaders['Share_ID']],
                                                        self._stockData[item][self._stockHeaders['Shareholder']],
                                                        self._stockData[item][self._stockHeaders['Quantity']],
                                                        "£ "+str(round(self._stockData[item][self._stockHeaders[
                                                            'Price_Per_Share']], 2)),
                                                        self._stockData[item][self._stockHeaders['Date']],
                                                        str(percentage)+"%"])
            quantityOfShares = self._stockHeaders['Quantity']
            self.boughtStocks = sum([elem[quantityOfShares] for elem in self._stockDataList])

        except Error as e:
            self._controller.log_event(e, self._controller.lineno())

        finally:
            self.conn.commit()
            self.conn.close()
            self.treeChild = self.stockTree.get_children('')

    def updateStockDetails(self):
        #function that gets the company's general stock data from the database
        try:
            self.conn = MySQLConnection(**self.ini_db)
            if self.conn.is_connected() != True:
                self._controller.log_event(self._controller.lineno(),
                                           "Database not connected contact the network admin.")
            self.c = self.conn.cursor()
            self.c.execute("SELECT * FROM stocks WHERE Company_Name=%(company)s", {'company':self.company[
                'Company_Name']})
            self._stockDetHeaders = {elem[1][0]:elem[0] for elem in enumerate(self.c.description)}
            self._stockDetData = self.c.fetchall()
            if self._stockDetData:
                self._stockDetData = self._stockDetData[0]

        except Error as e:
            self._controller.log_event(e, self._controller.lineno())
        finally:
            self.conn.commit()
            self.conn.close()

    def editStockDet(self):
        #navigate back to the frame to edit the company's general stock details
        self.setStockDetails(edit=True)
        self.Outstanding.set(self._stockDetData[self._stockDetHeaders['Total_Shares']])
        self.Price.set(self._stockDetData[self._stockDetHeaders['Share_Price']])

    def setStockDetails(self, edit=False):
        #creates the frame and widgets to set the company's general stock details
        self.stockDetFrame = tk.Frame(self, bg=self._bg)
        self.stockDetFrame.place(relx=0, rely=0, relheight=1, relwidth=1)
        self._head2 = tk.Label(self.stockDetFrame, text="Set Company Stock Details", bg=self._bg, font=self.fonts[0])
        self._head2.place(relx=0.3, rely=0.2, relheight=0.05, relwidth=0.4)

        self._outstandingL = tk.Label(self.stockDetFrame, text="Total Shares", bg=self._bg, font=self.fonts[1], anchor='e')
        self._outstandingL.place(relx=0.38, rely=0.28, relheight=0.05, relwidth=0.12)
        self.Outstanding = tk.IntVar()
        self._outstandingE = tk.Entry(self.stockDetFrame, textvariable=self.Outstanding, font=self.fonts[2],
                                      bg=self._bg)
        self._outstandingE.place(relx=0.5, rely=0.28, relheight=0.05, relwidth=0.12)

        self._priceL = tk.Label(self.stockDetFrame, text="Share Price", bg=self._bg, font=self.fonts[1],
                                      anchor='e')
        self._priceL.place(relx=0.38, rely=0.36, relheight=0.05, relwidth=0.12)
        self.Price = tk.DoubleVar()
        self._priceE = tk.Entry(self.stockDetFrame, textvariable=self.Price, font=self.fonts[2],
                                      bg=self._bg)
        self._priceE.place(relx=0.5, rely=0.36, relheight=0.05, relwidth=0.12)

        self._saveB = tk.Button(self.stockDetFrame, text="Save", font=self.fonts[2], bg=self._bg, command=lambda:
                                self.setStocks(edit))
        self._saveB.place(relx=0.54, rely=0.48, relheight=0.05, relwidth=0.07)

        if edit:
            self._backB = tk.Button(self.stockDetFrame, text="back", font=self.fonts[2], bg=self._bg,
                                    command=lambda: self.back(self.stockDetFrame))
            self._backB.place(relx=0.43, rely=0.48, relheight=0.05, relwidth=0.07)

    def setStocks(self, edit):
        #function validates inputs for stock details and updates the database
        try:
            float(self.Price.get())
            number, decimal = str(self.Price.get()).split(".")
            if len(decimal) > 2:
                float("wrong")

            self.Outstanding.get()
        except ValueError:
            tk.messagebox.showerror("Invalid Input", "Ensure the price is a number with no more than 2 decimal places and the total is an integer.")
            return

        if edit:
            if self.Outstanding.get() < self.boughtStocks:
                tk.messagebox.showerror("Invalid Input", "Your total shares cannot be below that the currently public held shares.")
                return

        try:
            #updates the database
            self.conn = MySQLConnection(**self.ini_db)
            if self.conn.is_connected() != True:
                self._controller.log_event(self._controller.lineno(),
                                           "Database not connected contact the network admin.")
            self.c = self.conn.cursor()

            marketCap = self.Outstanding.get() * self.Price.get()
            if edit:
                self.c.execute("UPDATE stocks SET Total_Shares=%(shares)s, Share_Price=%(price)s, "
                               "Market_Capitalisation=%(cap)s, Last_Updated=CURRENT_TIMESTAMP WHERE Company_Name=%("
                               "name)s", {'shares':self.Outstanding.get(), 'price':self.Price.get(), 'cap':marketCap,
                                'name':self.company['Company_Name']})
            else:
                self.c.execute("INSERT INTO stocks VALUES (%(name)s, %(shares)s, %(price)s, %(cap)s, "
                               "CURRENT_TIMESTAMP)",
                               {'shares':self.Outstanding.get(), 'price':self.Price.get(), 'cap':marketCap,
                                'name':self.company['Company_Name']})

        except Error as e:
            self._controller.log_event(e, self._controller.lineno())

        finally:
            self.conn.commit()
            self.conn.close()
            #refreshes the label to display the new stock details
            self.updateStockDetails()
            self._statsDetails.config(text=str(self._stockDetData[self._stockDetHeaders['Total_Shares']])+"\n\n"+str(
                self._stockDetData[self._stockDetHeaders['Share_Price']])+"\n\n"+str(self._stockDetData[
                self._stockDetHeaders['Market_Capitalisation']]))
            #returns back to the main stock page
            self.back(self.stockDetFrame)

    def createPie(self, hideLabels=False, window=False):
        #creates pie chart to show distribution of stocks
        if self.viewPie[0].get():
            if not window:
                #creates new frame to display the pie chart in the current window
                self.PieFrame = tk.LabelFrame(self, text="Stock Pie Chart", bg=self._bg)
                self.PieFrame.place(relx=0.01, rely=0.09, relheight=0.9, relwidth=0.72)
            else:
                #creates a new window to display the pie chart
                self.PieFrame = tk.Toplevel()
                self.PieFrame.geometry('700x700')
            index = self._stockHeaders['Shareholder']
            shareholders = [elem[index] for elem in self._stockDataList]
            index = self._stockHeaders['Quantity']
            shares = [elem[index] for elem in self._stockDataList]
            index = self._stockHeaders['Percent']
            percentages = [elem[index] for elem in self._stockDataList]

            #creates the slice for free shares and calculates the percentage
            sharesFree = int(self._stockDetData[self._stockDetHeaders['Total_Shares']]) - self.boughtStocks
            if sharesFree > 0:
                shareholders.append("Free")
                shares.append(sharesFree)
                percentages.append(100-sum(percentages))

            labels = ['{0} - {1} %'.format(a, b) for a, b in zip(shareholders, percentages)]

            #creates the plot
            fig = plt.Figure(figsize=(5, 4), dpi=100)
            fig1, ax1 = plt.subplots()
            explod = [0]*(len(shares)-1)
            explod.append(0.1)
            if not hideLabels:
                wedge, wedgeLabels, text = ax1.pie(shares, labels=shareholders, autopct='%1.1f%%', rotatelabels=True,
                                                  shadow=True,startangle=90, explode=explod, pctdistance=1.1,
                                                   labeldistance=1.2)
            else:
                wedge, text = ax1.pie(shares, rotatelabels=True,shadow=True,startangle=90, explode=explod,
                                      pctdistance=1.1, labeldistance=1.2)


            #sets chart attributes
            ax1.axis('equal')
            plt.legend(wedge, labels)
            plt.plot()

            #creates navigation toolbar
            _pieCanvas = FigureCanvasTkAgg(fig1, self.PieFrame)
            _pieCanvas.draw()
            _pieCanvas.get_tk_widget().place(relx=0, rely=0, relheight=1, relwidth=1)

            _pieToolbar = NavigationToolbar2Tk(_pieCanvas, self.PieFrame)
            _pieToolbar.update()
            _pieCanvas._tkcanvas.place(relx=0, rely=0, relheight=1, relwidth=1)
        else:
            #if checkbutton clicked again, destroy the pie chart
            self.PieFrame.destroy()
            plt.close('all')

    def downloadDB(self):
        #function downloads stocks to an excel spreadsheet
        self.style0 = easyxf("align:horiz centre")
        self.style1 = easyxf('align:horiz center', num_format_str='DD-MMM-YY')
        self.style2 = easyxf('pattern: pattern solid, fore_color yellow; font: bold on;align:horiz center')

        self.book = Workbook()

        self.sheet1 = self.book.add_sheet('Stocks')
        for CI, col in enumerate(self._stockHeadersList):
            self.sheet1.write(0, CI, col, self.style2)
        for r, row in enumerate(self._stockDataList):
            for c, col in enumerate(row):
                if isinstance(col, dt.date):
                    self.sheet1.write(r + 1, c, col, self.style1)
                else:
                    self.sheet1.write(r + 1, c, col, self.style0)

        self.file = filedialog.asksaveasfilename(title="Save Stock Data", initialdir='Log/',
                                                 defaultextension=".xls")
        if self.file:
            self.book.save(self.file)

    def back(self, frame=None):
        #function removes the list of frames, passed as an argument, from display and refreshes stock treeview
        if frame:
            frame.destroy()
        self.updateStocks()

class TransactionsPage(tk.Frame):
    def __init__(self, parent, controller, _bg):
        # creates stocks page, setting values for the root window, controlling class (Main), company details,
        # background colour, font styles, user details, user access level and ID and the general and company's private
        # database access details
        super().__init__(parent, bg=_bg)
        self.place(x=0, y=0, relheight=1, relwidth=1)
        self._bg = _bg
        self._top = parent
        self._controller = controller
        self.company = self._controller.get_company_metadata()
        self.user = controller.output_user()
        self.db = self._controller.getDBdetails()
        self.ini_db = self._controller.getInitDB()
        self.fonts = controller.getAppearance('font')

        #creates widgets to view transactions
        self._head1 = tk.Label(self, text="Transactions", bg=_bg, font=self.fonts[0])
        self._head1.place(relx=0.4, rely=0.01, relheight=0.05, relwidth=0.2)

        self._transHeaders = ("Transaction ID", "Details", "Quantity", "Individual Price","Total Value", "Type",
                              "Date")
        self.transTree = ttk.Treeview(self, height=20, columns=self._transHeaders, show='headings')
        for heading in self._transHeaders:
            self.transTree.heading(heading, text=heading, command=lambda
                col=heading: self._controller.sortItem(col, self.transTree))
            self.transTree.column(heading, stretch='yes', width=8)

        self.transTree.place(relx=0.05, rely=0.09, relheight=0.76, relwidth=0.68)

        self.treeScroll = ttk.Scrollbar(self, orient="vertical", command=self.transTree.yview)
        self.treeScroll.place(relx=0.73, rely=0.09, relheight=0.76)

        self.transTree.configure(yscrollcommand=self.treeScroll.set)

        self.transTree.bind("<Double-1>", lambda e:self.viewDetails())

        #search element searches the treeview
        self.Search_ = [tk.StringVar(), tk.StringVar()]
        self._searchL = tk.Label(self, text="Search", font=self.fonts[1], bg=_bg)
        self._searchL.place(relx=0.05, rely=0.91, relwidth=0.08, relheight=0.05)
        self._searchO = ttk.OptionMenu(self, self.Search_[0], '', *self._transHeaders)
        self._searchO.place(relx=0.13, rely=0.91, relwidth=0.09, relheight=0.05)
        self.Search_[0].set(self._transHeaders[0])
        self._searchE = tk.Entry(self, textvariable=self.Search_[1], bg=_bg, font=self.fonts[2])
        self._searchE.place(relx=0.22, rely=0.91, relwidth=0.13, relheight=0.05)
        self._searchB = tk.Button(self, text="Go", font=self.fonts[2], bg=_bg, command=lambda:
        self._controller.searchItem(
            self.Search_[0].get(), self.Search_[1].get(), self.transTree, self.treeChild), activebackground=_bg,
                                  highlightthickness=0, bd=0)
        self._searchB.place(relx=0.35, rely=0.91, relheight=0.05, relwidth=0.05)

        self._searchE.bind("<Return>", lambda e: self._controller.searchItem(
            self.Search_[0].get(), self.Search_[1].get(), self.transTree, self.treeChild))

        self._head3 = tk.Label(self, text="Overall Finances", bg=_bg, font=self.fonts[2])
        self._head3.place(relx=0.83, rely=0.2, relheight=0.05, relwidth=0.1)
        self._statsHeader = tk.Label(self, text="Income: \n\nExpenses: \n\nStocks Bought: \n\nOverall Value: ",
                                     font=self.fonts[2], bg=_bg, anchor='e', justify='right')
        self._statsHeader.place(relx=0.78, rely=0.26, relheight=0.2, relwidth=0.1)
        self._statsDetails = tk.Label(self, text="", font=self.fonts[2], bg=_bg, anchor='w', justify='left')
        self._statsDetails.place(relx=0.89, rely=0.26, relheight=0.2, relwidth=0.09)

        #button deletes transaction
        self.IMG1 = tk.PhotoImage(file='Images/Icons/bin.gif')
        self._delB = tk.Button(self, image=self.IMG1, bg=self._bg, bd=0, highlightthickness=0,
                               activebackground=self._bg, command=self.delTrans)
        self._delB.image = self.IMG1
        self._delB.place(relx=0.67, rely=0.91, relwidth=0.05, relheight=0.05)

        #button creates new transaction
        self.IMG2 = tk.PhotoImage(file='Images/Icons/add.gif')
        self._addB = tk.Button(self, image=self.IMG2, bg=self._bg, bd=0, highlightthickness=0,
                               activebackground=self._bg, command=self.addTrans)
        self._addB.image = self.IMG2
        self._addB.place(relx=0.84, rely=0.46, relwidth=0.07, relheight=0.07)

        #button refreshes treeview
        self.IMG3 = tk.PhotoImage(file='Images/Icons/reload.gif')
        self._reloadB = tk.Button(self, image=self.IMG3, bg=self._bg, bd=0,highlightthickness=0,
                               activebackground=self._bg, command=self.updateTrans)
        self._reloadB.image = self.IMG3
        self._reloadB.place(relx=0.75, rely=0.1, relwidth=0.04, relheight=0.05)

        #button downloads all transaction details to spreadsheet
        self.IMG4 = tk.PhotoImage(file='Images/Icons/download.gif')
        self._downloadB = tk.Button(self, image=self.IMG4, bg=self._bg, bd=0,highlightthickness=0,
                               activebackground=self._bg, command=self.downloadDB)
        self._downloadB.image = self.IMG4
        self._downloadB.place(relx=0.95, rely=0.2, relwidth=0.05, relheight=0.05)

        if self.user['Access_Level'] > 3:
            #disable the add, delete and download buttons if access level is greater than 3
            self._delB.config(state='disabled')
            self._addB.config(state='disabled')
            self._downloadB.config(state='disabled')

        #checkbutton to display graph of data
        self.viewChart = [tk.IntVar(), tk.StringVar(), tk.StringVar(), tk.IntVar()]
        self.IMG4 = tk.PhotoImage(file='Images/Icons/lineGraph.gif')
        self._graphB = tk.Checkbutton(self, image=self.IMG4, variable=self.viewChart[0], bd=0,
                                      highlightthickness=0, selectcolor='#ff8989', bg=_bg,
                                      indicatoron=False, command=lambda: self.createChart(
                res=self.viewChart[1].get(), scale=self.viewChart[2].get(), window=self.viewChart[
                3].get()))
        self._graphB.image = self.IMG4
        self._graphB.place(relx=0.8, rely=0.58, relwidth=0.17, relheight=0.18)

        self._resL = tk.Label(self, text="Resolution", bg=_bg, font=self.fonts[2])
        self._resL.place(relx=0.8, rely=0.83, relheight=0.04, relwidth=0.06)
        self._xAxisIncrL = tk.Label(self, text="Scale", bg=_bg, font=self.fonts[2])
        self._xAxisIncrL.place(relx=0.9, rely=0.83, relheight=0.04, relwidth=0.03)
        #select how detailed the graph is
        self._resE = ttk.OptionMenu(self, self.viewChart[1], *('Daily', 'Hourly', 'Daily', 'Weekly', 'Monthly'))
        self._resE.place(relx=0.81, rely=0.88, relheight=0.05, relwidth=0.06)
        #select the x-axis intervals
        self._xAxisIncrE = ttk.OptionMenu(self, self.viewChart[2], *('Daily', 'Hourly', 'Daily',
                                                                     'Weekly', 'Monthly', 'Yearly'))
        self._xAxisIncrE.place(relx=0.89, rely=0.88, relheight=0.05, relwidth=0.06)

        #open graph in new window
        self._windowB = tk.Checkbutton(self, text="New window", variable=self.viewChart[3], bg=self._bg)
        self._windowB.place(relx=0.84, rely=0.93, relheight=0.05, relwidth=0.1)
        #update treeview
        self.updateTrans()

    def addTrans(self):
        #function creates frame to add transaction details
        self.addFrame = tk.Frame(self, bg=self._bg)
        self.addFrame.place(relx=0, rely=0, relheight=1, relwidth=1)
        self._title2 = tk.Label(self.addFrame, text="Add transaction", bg=self._bg, font=self.fonts[0])
        self._title2.place(relx=0.45, rely=0.05, relheight=0.05, relwidth=0.2)
        # Labels
        self._idL = tk.Label(self.addFrame, text="Transaction ID (if exists)", bg=self._bg, font=self.fonts[1],
                             anchor='e')
        self._idL.place(relx=0.28, rely=0.15, relheight=0.05, relwidth=0.2)

        self._detL = tk.Label(self.addFrame, text="Details", bg=self._bg, font=self.fonts[1], anchor='e')
        self._detL.place(relx=0.28, rely=0.25, relheight=0.05, relwidth=0.2)

        self._descriptionL = tk.Label(self.addFrame, text="Description", bg=self._bg, font=self.fonts[1], anchor='e')
        self._descriptionL.place(relx=0.28, rely=0.35, relheight=0.05, relwidth=0.2)

        self._quantityL = tk.Label(self.addFrame, text="Quantity", bg=self._bg, font=self.fonts[1], anchor='e')
        self._quantityL.place(relx=0.28, rely=0.55, relheight=0.05, relwidth=0.2)

        self._pricepaidL = tk.Label(self.addFrame, text="Price Paid (per unit)", bg=self._bg, font=self.fonts[1],
                                    anchor='e')
        self._pricepaidL.place(relx=0.28, rely=0.65, relheight=0.05, relwidth=0.2)

        self._typeL = tk.Label(self.addFrame, text="Type", bg=self._bg, font=self.fonts[1], anchor='e')
        self._typeL.place(relx=0.28, rely=0.75, relheight=0.05, relwidth=0.2)

        # entry widgets
        self._ID = tk.StringVar()
        self._idE = tk.Entry(self.addFrame, textvariable=self._ID, bg=self._bg)
        self._idE.place(relx=0.5, rely=0.15, relheight=0.05, relwidth=0.2)
        self._loadB = tk.Button(self.addFrame, text="Load", bg=self._bg, font=self.fonts[2],
                                command=self.loadTrans)
        self._loadB.place(relx=0.7, rely=0.15, relheight=0.05, relwidth=0.05)

        self._Det = tk.StringVar()
        self._detE = tk.Entry(self.addFrame, textvariable=self._Det, bg=self._bg)
        self._detE.place(relx=0.5, rely=0.25, relheight=0.05, relwidth=0.2)

        self._descriptionE = tk.scrolledtext.ScrolledText(master=self.addFrame, wrap=tk.WORD, bg=self._bg)
        self._descriptionE.place(relx=0.5, rely=0.35, relheight=0.15, relwidth=0.2)

        self._Quantity = tk.IntVar()
        self._quantityE = tk.Entry(self.addFrame, textvariable=self._Quantity, bg=self._bg)
        self._quantityE.place(relx=0.5, rely=0.55, relheight=0.05, relwidth=0.2)

        self._Price = tk.DoubleVar()
        self._pricepaidE = tk.Entry(self.addFrame, textvariable=self._Price, bg=self._bg)
        self._pricepaidE.place(relx=0.5, rely=0.65, relheight=0.05, relwidth=0.2)

        self._Type = tk.StringVar()
        self._typeE = ttk.OptionMenu(self.addFrame, self._Type, *('', 'income', 'expenditure'))
        self._Type.set("income")
        self._typeE.place(relx=0.5, rely=.75, relheight=0.05, relwidth=0.2)

        #returns to the main transaction page
        self._cancelB = tk.Button(self.addFrame, text="Cancel", bg=self._bg, font=self.fonts[2],
                                  command=lambda: self.back(self.addFrame))
        self._cancelB.place(relx=0.43, rely=0.87, relheight=0.05, relwidth=0.1)
        #saves transaction details
        self._nextB = tk.Button(self.addFrame, text="Apply", bg=self._bg, font=self.fonts[2], command=self.add)
        self._nextB.place(relx=0.59, rely=0.87, relheight=0.05, relwidth=0.1)

    def add(self):
        #validate and add the input transaction to the database
        if not 0 < len(self._Det.get()) <= 16777215:
            tk.messagebox.showerror("Invalid Input", "Ensure to enter some title/details for the details field less "
                                                     "than 16,777,215 characters.")
            return
        if not 0 < len(self._descriptionE.get(1.0, 'end')) <= 16777215:
            tk.messagebox.showerror('Invalid Input',
                                    'Enter something in the description field (must not exceed 16,777,215 '
                                    'characters).')
            return

        try:
            self._Quantity.get()
            float(self._Price.get())
            number, decimal = str(self._Price.get()).split('.')
            if len(decimal) > 2:
                float('no')
            if self._Price.get() < 0:
                float("no")
        except ValueError:
            tk.messagebox.showerror("Invalid Input", "Ensure the quantity is an integer and the price is a positive "
                                                     "decimal.")
            return


        try:
            #add transaction to database
            self.conn = MySQLConnection(**self.db)
            if self.conn.is_connected() != True:
                self._controller.log_event(self._controller.lineno(),
                                           "Database not connected contact the network admin.")
            self.c = self.conn.cursor()
            uniqueID = self._controller.getID()
            total = self._Price.get()*self._Quantity.get()
            self.c.execute("INSERT INTO income_expenditure VALUES (%(id)s, %(det)s, %(descrip)s, %(price)s, "
                           "%(quantity)s, %(total)s, %(type)s, CURRENT_TIMESTAMP, NULL)", {'id':uniqueID,
                            'det':self._Det.get(), 'descrip':self._descriptionE.get(1.0, 'end'),
                            'price':self._Price.get(), 'quantity':self._Quantity.get(), 'total':round(total,2),
                            'type':self._Type.get()})
            tk.messagebox.showinfo("Success", "Transaction successfully recorded!")
        except Error as e:
            self._controller.log_event(e, self._controller.lineno())
        finally:
            self.conn.commit()
            self.conn.close()
            self.updateTrans()
            self.back(self.addFrame)

    def loadTrans(self):
        #load a transactions details and populate the add transaction's fields with it when a transaction id is input
        # and loaded from
        if self._ID.get().isnumeric() and self._ID.get() in self._transData:
            selectionData = self._transData[self._ID.get()]
            self._Det.set(selectionData[self._transHeaders['Details']])
            self._descriptionE.delete(1.0, 'end')
            self._descriptionE.insert(1.0, selectionData[self._transHeaders['Description']])
            self._Quantity.set(str(selectionData[self._transHeaders['Quantity']]))
            self._Price.set(str(selectionData[self._transHeaders['Price_Per_Piece']]))
            self._Type.set(selectionData[self._transHeaders['Type']])
        else:
            tk.messagebox.showerror("Invalid Input", "The transaction ID entered was invalid.")

    def delTrans(self):
        #delete created transaction from the database
        if not self.transTree.selection():
            #checks if item from treeview is selected
            tk.messagebox.showerror("Select transaction", "Please select the transaction you wish to delete.")
            return

        #this page can only delete generic transactions, if a stock/item/sale transaction selected, program alerts to
        # navigate to corresponding page
        self.selectionData = self.transTree.item(self.transTree.selection(), 'values')
        if self.selectionData[5] == 'income - Sales' or self.selectionData[5] == 'income - Shares':
            tk.messagebox.showerror("Invalid transaction", "To delete shares or sales please go to their respective "
                                                           "pages. Please choose an income/expenditure transaction to delete.")
            return
        elif self._transData[self.selectionData[0]][self._transHeaders['Item_Batch_ID']]:
            tk.messagebox.showerror("Invalid transaction", "Records say this transaction was from buying an item and "
                                                           "registering it in the inventory. If you want to delete "
                                                           "this item go to the Inventory menu.")
            return

        #deletes transaction from database
        try:
            self.conn = MySQLConnection(**self.db)
            if self.conn.is_connected() != True:
                self._controller.log_event(self._controller.lineno(),
                                           "Database not connected contact the network admin.")
            self.c = self.conn.cursor()
            self.c.execute("DELETE FROM income_expenditure WHERE Transaction_ID=%(id)s", {'id':self.selectionData[0]})
            tk.messagebox.showinfo("Success", "Transaction deleted successfully!")
        except Error as e:
            self._controller.log_event(e, self._controller.lineno())
        finally:
            self.conn.commit()
            self.conn.close()
            #refresh treeview
            self.updateTrans()

    def viewDetails(self):
        #function displays the transaction's details in a pop-up window
        if not self.transTree.selection():
            #checks if transaction is selected
            tk.messagebox.showerror("Select item", "Please select a transaction.")
            return
        treeSelecData = self.transTree.item(self.transTree.selection(), 'values')
        self.selectionID = treeSelecData[0]
        if treeSelecData[5] == 'income - Sales':
            selectionData = self._salesData[self.selectionID]
            selectionHeaders = self._salesHeadersList
            title = 'Sales ID: '+self.selectionID
        elif treeSelecData[5] == 'income - Shares':
            selectionData = self._stockData[int(self.selectionID)]
            selectionHeaders = self._stockHeadersList
            title = 'Shares ID: '+self.selectionID
        else:
            selectionData = self._transData[self.selectionID]
            selectionHeaders = self._transHeadersList
            title = 'Transaction ID: '+self.selectionID

        details = []
        self._top.clipboard_clear()
        #copy the unique ID of the particular transaction to the systems clipboard
        self._top.clipboard_append(self.selectionID)
        for a, b in zip(selectionHeaders, selectionData):
            details.append("{0}: {1}\n".format(a, b))
        tk.messagebox.showinfo(title,"".join(details))

    def updateTransDet(self):
        #update the general finance details in the label widget
        stocks = sum([round(self._stockData[elem][self._stockHeaders['Quantity']]*self._stockData[elem]
                    [self._stockHeaders['Price_Per_Share']],2) for elem in self._stockData])
        sales = sum([self._salesData[elem][self._salesHeaders['Total_Income']] for elem in self._salesData])
        transIncome = list(filter(lambda elem: self._transData[elem][self._transHeaders['Type']]
                                                                    == 'income', self._transData))
        transIncome = sum([self._transData[elem][self._transHeaders['Total_Cost']] for elem in transIncome])
        transSpend = list(filter(lambda elem: self._transData[elem][self._transHeaders['Type']] == 'expenditure',
                                 self._transData))
        transSpend = sum([self._transData[elem][self._transHeaders['Total_Cost']] for elem in transSpend])

        income = transIncome + sales + stocks
        overall = income - transSpend
        self._statsDetails.config(text="£ %.2f \n\n£ %.2f\n\n £ %.2f\n\n £ %.2f" % (income, transSpend, stocks,
                                                                                    overall))

    def updateTrans(self):
        #clear the treeview and re-populate it with the latest transaction data from the database
        self.transTree.delete(*self.transTree.get_children())
        try:
            self.conn = MySQLConnection(**self.db)
            if self.conn.is_connected() != True:
                self._controller.log_event(self._controller.lineno(),
                                           "Database not connected contact the network admin.")
            self.c = self.conn.cursor()
            self.c.execute("SELECT * FROM income_expenditure")

            self._transHeadersList = [elem[0] for elem in self.c.description]
            self._transHeaders = {elem[1]:elem[0] for elem in enumerate(self._transHeadersList)}
            self._transData = {elem[0]:list(elem) for elem in self.c.fetchall()}

            for index, item in enumerate(self._transData):
                self.transTree.insert("", "end", values=[self._transData[item][self._transHeaders['Transaction_ID']],
                                                        self._transData[item][self._transHeaders['Details']],
                                                        self._transData[item][self._transHeaders['Quantity']],
                                                        "£ "+str(round(self._transData[item][self._transHeaders[
                        'Price_Per_Piece']],2)),
                                                        "£ "+str(round(self._transData[item][self._transHeaders[
                        'Total_Cost']],2)),
                                                        self._transData[item][self._transHeaders['Type']],
                                                        self._transData[item][self._transHeaders['Date']]])


            self.c.execute("SELECT * FROM stock")
            self._stockHeadersList = [elem[0] for elem in self.c.description]
            self._stockHeaders = {elem[1]:elem[0] for elem in enumerate(self._stockHeadersList)}
            self._stockData = {elem[0]:list(elem) for elem in self.c.fetchall()}
            for index, item in enumerate(self._stockData):
                self.transTree.insert("", "end", values=[self._stockData[item][self._stockHeaders['Share_ID']],
                                                        self._stockData[item][self._stockHeaders['Shareholder']],
                                                        self._stockData[item][self._stockHeaders['Quantity']],
                                                        "£ "+str(round(self._stockData[item][self._stockHeaders[
                        'Price_Per_Share']],2)),
                                                        "£ "+str(round(self._stockData[item][self._stockHeaders[
                        'Price_Per_Share']]*self._stockData[item][self._stockHeaders['Quantity']], 2)),
                                                        "income - Shares",
                                                        self._stockData[item][self._stockHeaders['Date']]])


            self.c.execute("SELECT sales.*, inventory.* FROM sales "
                           "INNER JOIN all_inventory "
                           "ON sales.Item_Batch_ID = all_inventory.Batch_ID "
                           "INNER JOIN inventory "
                           "ON all_inventory.Item_ID = inventory.Item_ID")
            self._salesHeadersList = [elem[0] for elem in self.c.description]
            self._salesHeaders = {elem[1]:elem[0] for elem in enumerate(self._salesHeadersList)}
            self._salesData = {elem[0]:list(elem) for elem in self.c.fetchall()}
            for index, item in enumerate(self._salesData):
                self.transTree.insert("", "end", values=[self._salesData[item][self._salesHeaders['Sale_ID']],
                                                        self._salesData[item][self._salesHeaders['Item_Name']],
                                                        self._salesData[item][self._salesHeaders['Quantity']],
                                                        "£ "+str(round(self._salesData[item][self._salesHeaders[
                        'Price_Sold']],2)),
                                                        "£ "+str(round(self._salesData[item][self._salesHeaders[
                        'Total_Income']], 2)),
                                                        "income - Sales",
                                                        self._salesData[item][self._salesHeaders['Date']]])

            self._controller.sortItem('Date', self.transTree, alternateSort=False)
        except Error as e:
            self._controller.log_event(e, self._controller.lineno())

        finally:
            self.conn.commit()
            self.conn.close()
            self.treeChild = self.transTree.get_children('')
            self.updateTransDet()

    def createDate(self, *date):
        #convert datetime objects to matplotlib date/time objects
        mplDt = []
        for x in date:
            mplDt.append(date2num(x))
        return mplDt

    def createChart(self, res, scale, window=False):
        #creates the graph of the company's finances
        if (len(self._salesData) + len(self._stockData) + len(self._transData)) <= 1:
            tk.messagebox.showerror("Insufficient data", "There is insufficient data to plot a graph. Please ensure there are more than 1 transaction entries.")
            return
        if self.viewChart[0].get():
            if not window:
                #creates a frame in the current window
                self.ChartFrame = tk.LabelFrame(self, text="Transactions Pie Chart", bg=self._bg)
                self.ChartFrame.place(relx=0.01, rely=0.07, relheight=0.92, relwidth=0.74)
                self.windowedChart = False
            else:
                #creates a new window to display the graph
                self.ChartFrame = tk.Toplevel()
                self.ChartFrame.geometry('700x700')
                self.windowedChart = True

            #gets the plottable points for the income and expenditure line
            incomePoints = []
            for item in self._transData:
                if self._transData[item][self._transHeaders['Type']] == 'income':
                    total = self._transData[item][self._transHeaders['Total_Cost']]
                    date = self._transData[item][self._transHeaders['Date']]
                    incomePoints.append([date, total])
            for item in self._stockData:
                total = round(self._stockData[item][self._stockHeaders['Quantity']]*self._stockData[item][
                    self._stockHeaders['Price_Per_Share']], 2)
                date = self._stockData[item][self._stockHeaders['Date']]
                incomePoints.append([date, total])
            for item in self._salesData:
                total = self._salesData[item][self._salesHeaders['Total_Income']]
                date = self._salesData[item][self._salesHeaders['Date']]
                incomePoints.append([date, total])

            incomePoints.sort()

            expensePoints = []
            for item in self._transData:
                if self._transData[item][self._transHeaders['Type']] == 'expenditure':
                    date = self._transData[item][self._transHeaders['Date']]
                    total = self._transData[item][self._transHeaders['Total_Cost']]
                    expensePoints.append([date, total])

            expensePoints.sort()

            #create x axis values
            xAxisValues = []
            incomeDates = [elem[0] for elem in incomePoints]
            expenseDates = [elem[0] for elem in expensePoints]
            minim = [min(incomeDates), min(expenseDates)]
            xAxisValues.append(min(minim))
            xAxisValues[0] = xAxisValues[0].replace(minute=0, second=0, microsecond=0)
            maxim = [max(incomeDates), max(expenseDates)]
            xAxisValues.append(max(maxim))
            xAxisValues[1] = xAxisValues[1].replace(minute=0, second=0, microsecond=0)
            if res == 'Hourly':
                res = 1
            elif res == 'Daily':
                res = 24
            elif res == 'Weekly':
                res = 168
            elif res == 'Monthly':
                res = 720
            xAxisValues[1] += dt.timedelta(hours=res)
            xAxisValues.insert(0, xAxisValues[0])
            while xAxisValues[-2] != xAxisValues[-1]:
                xAxisValues.insert(-1, xAxisValues[-2]+dt.timedelta(hours=1))
            del xAxisValues[0]


            incomePlots = []
            total = 0
            current = 1
            for item in incomePoints:
                while True:
                    if item[0] <= xAxisValues[current]:
                        total += item[1]
                        break
                    else:
                        incomePlots.append([self.createDate(xAxisValues[current]), total])
                        current+=1
            incomePlots.append([self.createDate(xAxisValues[current]), total])
            for x in xAxisValues[len(incomePlots):]:
                incomePlots.append([x, incomePlots[-1][1]])

            expensePlots = []
            total = 0
            current = 1
            for item in expensePoints:
                while True:
                    if item[0] <= xAxisValues[current]:
                        total += item[1]
                        break
                    else:
                        expensePlots.append([self.createDate(xAxisValues[current]), total])
                        current+=1
            expensePlots.append([self.createDate(xAxisValues[current]), total])
            for x in xAxisValues[len(expensePlots):]:
                expensePlots.append([x, expensePlots[-1][1]])

            #converts list of x-axis values to array
            xAxisValues = NPArray(self.createDate(*xAxisValues))
            #sets x-axis intervals and alerts user if it is too small
            timeDelta = num2date(xAxisValues.max()) - num2date(xAxisValues.min())
            if scale == 'Hourly':
                rule = rrulewrapper(HOURLY, interval=1)
                if timeDelta > dt.timedelta(days=20):
                    self.ChartFrame.destroy()
                    self._graphB.toggle()
                    tk.messagebox.showerror("Interval too short", "This interval is too short to display all the data "
                                                                  "since this data set currently spans a date range of "
                                                                  "greater than 20 days. Please choose a greater interval.")
                    return
            elif scale == 'Daily':
                rule = rrulewrapper(DAILY, interval=1)
                if timeDelta > dt.timedelta(days=365):
                    self.ChartFrame.destroy()
                    self._graphB.toggle()
                    tk.messagebox.showerror("Interval too short", "This interval is too short to display all the data "
                                                                  "since this data set currently spans a date range of "
                                                                  "greater than 365 days. Please choose a greater "
                                                                  "interval.")
                    return
            elif scale == 'Weekly':
                rule = rrulewrapper(WEEKLY, interval=1)
                if timeDelta > dt.timedelta(days=3360):
                    self.ChartFrame.destroy()
                    self._graphB.toggle()
                    tk.messagebox.showerror("Interval too short", "This interval is too short to display all the data "
                                                                  "since this data set currently spans a date range of "
                                                                  "greater than 3360 days. Please choose a greater "
                                                                  "interval.")
                    return
            elif scale == 'Monthly':
                rule = rrulewrapper(MONTHLY, interval=1)
                if timeDelta > dt.timedelta(days=14400):
                    self.ChartFrame.destroy()
                    self._graphB.toggle()
                    tk.messagebox.showerror("Interval too short", "This interval is too short to display all the data "
                                                                  "since this data set currently spans a date range of "
                                                                  "greater than 14400 days. Please choose a greater "
                                                                  "interval.")
                    return
            elif scale == 'Yearly':
                rule = rrulewrapper(YEARLY, interval=1)
                if timeDelta > dt.timedelta(days=175200):
                    self.ChartFrame.destroy()
                    self._graphB.toggle()
                    tk.messagebox.showerror("Interval too short", "This interval is too short to display all the data "
                                                                  "since this data set currently spans a date range of "
                                                                  "greater than 175200 days. Please choose a greater "
                                                                  "interval.")
                    return

            #create plot
            fig = plt.Figure(figsize=(5, 4), dpi=200)
            fig1, ax = plt.subplots()
            plt.subplots_adjust(bottom=0.13)
            line1 = NPArray([elem[1] for elem in incomePlots])
            line2 = NPArray([elem[1] for elem in expensePlots])

            ax.plot_date(xAxisValues, line1, fmt='g-', label='Income')
            ax.plot_date(xAxisValues, line2, fmt='r-', label='Expenditure')

            ax.fill_between(xAxisValues, line1, line2, where=(line1 > line2), facecolor='green',
                            interpolate=True)
            ax.fill_between(xAxisValues, line1, line2, where=(line1 < line2), facecolor='red',
                            interpolate=True)

            #sets graph attributes
            ax.axis('tight')
            ax.grid(color='g', linestyle=':')
            ax.xaxis_date()
            loc = RRuleLocator(rule)
            formatter = DateFormatter("%d '%b (%H:%M)")

            ax.xaxis.set_major_locator(loc)
            ax.xaxis.set_major_formatter(formatter)
            labelsx = ax.get_xticklabels()
            plt.setp(labelsx, rotation=30, fontsize=7)

            #sets axis, title and legend labels
            plt.xlabel("Date")
            plt.ylabel("Amount /£")
            plt.legend()
            plt.suptitle("Cumulative Income and Expenditure")

            plt.grid(True)

            #creates the navigation toolbar
            self._chartCanvas = FigureCanvasTkAgg(fig1, self.ChartFrame)
            self._chartCanvas.draw()
            self._chartCanvas.get_tk_widget().place(relx=0, rely=0, relheight=0.94, relwidth=1)

            self._chartBar = NavigationToolbar2Tk(self._chartCanvas, self.ChartFrame)
            self._chartBar.update()
            self._chartCanvas._tkcanvas.place(relx=0, rely=0, relheight=0.94, relwidth=1)

        else:
            #if checkbutton clicked off, destroy the graph and encosing frame/window
            self._chartCanvas.get_tk_widget().destroy()
            self.ChartFrame.destroy()

    def downloadDB(self):
        #download all transaction details to separate sheets in an excel spreadsheet
        self.style0 = easyxf("align:horiz centre")
        self.style1 = easyxf('align:horiz center', num_format_str='DD-MMM-YY')
        self.style2 = easyxf('pattern: pattern solid, fore_color yellow; font: bold on;align:horiz center')

        self.book = Workbook()

        sheet = self.book.add_sheet('Transactions')
        for CI, col in enumerate(self._transHeadersList):
            sheet.write(0, CI, col, self.style2)
        for r, row in enumerate(list(self._transData.values())):
            for c, col in enumerate(row):
                if isinstance(col, dt.date):
                    sheet.write(r + 1, c, col, self.style1)
                else:
                    sheet.write(r + 1, c, col, self.style0)
        sheet = self.book.add_sheet('Stocks')
        for CI, col in enumerate(self._stockHeadersList):
            sheet.write(0, CI, col, self.style2)
        for r, row in enumerate(list(self._stockData.values())):
            for c, col in enumerate(row):
                if isinstance(col, dt.date):
                    sheet.write(r + 1, c, col, self.style1)
                else:
                    sheet.write(r + 1, c, col, self.style0)

        sheet = self.book.add_sheet('Sales')
        for CI, col in enumerate(self._salesHeadersList):
            sheet.write(0, CI, col, self.style2)
        for r, row in enumerate(list(self._salesData.values())):
            for c, col in enumerate(row):
                if isinstance(col, dt.date):
                    sheet.write(r + 1, c, col, self.style1)
                else:
                    sheet.write(r + 1, c, col, self.style0)

        self.file = filedialog.asksaveasfilename(title="Save Stock Data", initialdir='Log/',
                                                 defaultextension=".xls")
        if self.file:
            self.book.save(self.file)

    def back(self, frame=None):
        #detroy the passed argument frame from display and update the treeview with the latest data
        if frame:
            frame.destroy()
        self.updateTrans()

class BuyInventoryPage(tk.Frame):
    def __init__(self, parent, controller, _bg):
        #creates event page, setting values for the root window, controlling class (Main), company details,
        # background colour, font styles, user details, user access level and ID, ebay api key and the company's
        # private database access details
        super().__init__(parent, bg=_bg)
        self.place(x=0, y=0, relheight=1, relwidth=1)
        self._bg = _bg
        self._top = parent
        self._controller = controller
        self.company = self._controller.get_company_metadata()
        self.user = controller.output_user()
        self.db = self._controller.getDBdetails()
        self.ebayID = self._controller.getEbayDetails()
        self.fonts = controller.getAppearance('font')
        self.userID = controller.output_user()['ID']

        #get current date
        self.date = dt.datetime.now()

        #create widgets to display current inventory
        self._title = tk.Label(self, text="Current Inventory", bg=_bg, font=self.fonts[0])
        self._title.place(relx=0.4, rely=0.01, relheight=0.05, relwidth=0.2)

        #sidebutton shows more functions to edit item
        self._sideButton = tk.Button(self, text=":::", font=self.fonts[0], command=self.openSideBar, bg=_bg)
        self._sideButton.place(relx=0.01, rely=0.05, relheight=0.05, relwidth=0.05)

        #binding opens sidebar when mouse hovers over sidebar
        self._sideButton.bind("<Enter>", lambda e:self.openSideBar())

        self._treeHeaders = ('Item ID', 'Item Name', 'Price Selling', 'Stock', 'Last Refresh')
        #treeview displays all current inventory
        self.itemTree = ttk.Treeview(self, height=20, columns=self._treeHeaders, show='headings')
        for heading in self._treeHeaders:
            self.itemTree.heading(heading, text=heading, command=lambda
                col=heading: self._controller.sortItem(col, self.itemTree))
            self.itemTree.column(heading, stretch='yes', width=8)

        self.itemTree.place(relx=0.05, rely=0.15, relheight=0.7, relwidth=0.65)

        self.treeScroll = ttk.Scrollbar(self, orient="vertical", command=self.itemTree.yview)
        self.treeScroll.place(relx=0.7, rely=0.15, relheight=0.7)

        self.itemTree.configure(yscrollcommand=self.treeScroll.set)
        #when selecting an item in treeview, label of item details to the right of the page
        self.itemTree.bind("<<TreeviewSelect>>", lambda e:self.viewItemDetails())

        self._itemTitle = tk.Label(self, text="Item Details", bg=_bg, font=self.fonts[1])
        self._itemTitle.place(relx=0.75, rely=0.1, relheight=0.05, relwidth=0.15)

        #displays YE logo image
        self.IMG = tk.PhotoImage(file='Images/young enterprise 2.gif')
        self._Logo = tk.Label(self, image=self.IMG, bg=_bg)
        self._Logo.image = self.IMG
        self._Logo.place(relx=0.75, rely=0.15, relheight=0.2, relwidth=0.2)


        #label widgets to display item stats when item in treeview clicked
        self.statsHeader = tk.Label(self, text="Item ID: \n\nItem Name: \n\nDate Added: \n\nCategory: "
                                               "\n\nCurrent Batch: \n\nOriginal "
                                               "Stock: \n\nOn Offer: \n\nOffers: \n\nEbay "
                                               "Price: ",
                                    bg=_bg, font=self.fonts[2], justify='left')
        self.statsHeader.place(relx=0.75, rely=0.4, relheight=0.4, relwidth=0.1)
        self.statsField = tk.Label(self, text="", bg=_bg, font=self.fonts[2], justify='right')
        self.statsField.place(relx=0.85, rely=0.4, relheight=0.4, relwidth=0.14)

        #allows sell, restock and discarding of items in current inventory
        self.XFunctions = tk.StringVar()
        self.XFunctions.set("Sell")
        self.Quantity = tk.IntVar()

        self.ExtraM = ttk.OptionMenu(self, self.XFunctions, *('','Sell', 'Restock', 'Discard Item'))
        self.ExtraM.place(relx=0.75, rely=0.85, relheight=0.05, relwidth=0.1)

        self.QuantEntry = tk.Entry(self, textvariable=self.Quantity, bg=_bg, font=self.fonts[2])
        self.QuantButton = tk.Button(self, text="Save", bg=_bg, font=self.fonts[2], command=self.xtraFunctions)
        self.QuantEntry.place(relx=0.86, rely=0.85, relheight=0.05, relwidth=0.06)
        self.QuantButton.place(relx=0.93, rely=0.85, relheight=0.05, relwidth=0.05)

        self.QuantEntry.bind("<Return>", lambda e:self.xtraFunctions())

        self.createSideBar()
        #close side bar when mouse no longer hovers over it
        self.sidebar.bind("<Leave>", lambda e:self.sidebar.place_forget())

        #delete created inventory item
        self.IMG1 = tk.PhotoImage(file='Images/Icons/bin.gif')
        self._delB = tk.Button(self, image=self.IMG1, bg=self._bg, bd=0,highlightthickness=0,
                               activebackground=self._bg, command=self.delItem)
        self._delB.image = self.IMG1
        self._delB.place(relx=0.63, rely=0.91, relwidth=0.05, relheight=0.05)

        #buy item with paypal
        self.IMG2 = tk.PhotoImage(file='Images/Icons/paypal.gif')
        self._paypalB = tk.Button(self, image=self.IMG2, bg=self._bg, bd=0,highlightthickness=0,
                               activebackground=self._bg, command=lambda: self.sellItem(paypal=True))
        self._paypalB.image = self.IMG2
        self._paypalB.place(relx=0.86, rely=0.9, relwidth=0.07, relheight=0.1)

        #search element, searches the treeview
        self.Search_ = [tk.StringVar(), tk.StringVar()]
        self._searchL = tk.Label(self, text="Search", font=self.fonts[1], bg=_bg)
        self._searchL.place(relx=0.05, rely=0.91, relwidth=0.08, relheight=0.05)
        self._searchO = ttk.OptionMenu(self, self.Search_[0], '',*self._treeHeaders)
        self._searchO.place(relx=0.13, rely=0.91, relwidth=0.09, relheight=0.05)
        self.Search_[0].set(self._treeHeaders[0])
        self._searchE = tk.Entry(self, textvariable=self.Search_[1], bg=_bg, font=self.fonts[2])
        self._searchE.place(relx=0.22, rely=0.91, relwidth=0.13, relheight=0.05)
        self._searchB = tk.Button(self, text="Go", font=self.fonts[2], bg=_bg, command=lambda:
        self._controller.searchItem(
            self.Search_[0].get(), self.Search_[1].get(), self.itemTree, self.treeChild), activebackground=_bg,
                                  highlightthickness=0, bd=0)
        self._searchB.place(relx=0.35, rely=0.91, relheight=0.05, relwidth=0.05)

        self._searchE.bind("<Return>", lambda e: self._controller.searchItem(
            self.Search_[0].get(), self.Search_[1].get(), self.itemTree, self.treeChild))

        #refresh treeview
        self.updateItemList()

    def delItem(self):
        #function completely deletes item from database
        if not self.itemTree.selection():
            #checks if item is selected
            tk.messagebox.showinfo("Select item", "Please select an item you wish to delete.")
            return
        self.selectionID = self.itemTree.item(self.itemTree.selection(), 'values')[0]
        if tk.messagebox.askyesno("Confirm", "Are you sure you wish to delete this product? All records about it ("
                                             "including sales) will be deleted."):
            #if confirmation still True, delete all records of item from everywhere in database
            try:
                self.conn = MySQLConnection(**self.db)
                if self.conn.is_connected() != True:
                    self._controller.log_event(self._controller.lineno(), "Database not connected contact the network "
                                                                          "admin.")
                self.c = self.conn.cursor()
                self.c.execute("DELETE FROM current_inventory WHERE Item_ID=%(id)s", {'id':self.selectionID})
                self.c.execute("DELETE FROM inventory WHERE Item_ID=%(id)s", {'id':self.selectionID})
                self.c.execute("DELETE FROM offers WHERE Item_ID=%(id)s", {'id':self.selectionID})
                self.c.execute("SELECT Batch_ID FROM all_inventory WHERE Item_ID=%(id)s AND Date_Removed IS NULL",
                               {'id':self.selectionID})
                self._batchID = [elem[0] for elem in self.c.fetchall()]
                self.c.execute("SELECT Item_ID, COUNT(*) FROM all_inventory WHERE Item_ID=%(id)s GROUP BY Item_ID",
                               {'id':self.selectionID})
                self._IDCount = self.c.fetchall()[0][0]
                self.c.execute("DELETE FROM all_inventory WHERE Item_ID=%(id)s AND Date_Removed IS NULL",
                               {'id':self.selectionID})
                for id in self._batchID:
                    self.c.execute("DELETE FROM income_expenditure WHERE Item_Batch_ID=%(id)s",
                                   {'id':id})
                    self.c.execute("DELETE FROM sales WHERE Item_Batch_ID=%(id)s",
                                   {'id':id})

                if self._IDCount == 1:
                    self.c.execute("DELETE FROM ebay_search_data WHERE Item_ID=%(id)s", {'id':self.selectionID})
                    self.c.execute("DELETE FROM ebay_price_history WHERE Item_ID=%(id)s", {'id':self.selectionID})

            except Error as e:
                self._controller.log_event(e, self._controller.lineno())
            finally:
                self.conn.commit()
                self.conn.close()

            #remove the item from the treeview and dictionary
            self.itemTree.delete(self.itemTree.selection())
            del self._itemData[int(self.selectionID)]

    def viewItemDetails(self):
        #function edits the stats label on the right to view the selected item's details
        if not self.itemTree.selection():
            tk.messagebox.showinfo("Select item", "Please select an item.")
            return
        self.selectionID = self.itemTree.item(self.itemTree.selection(), 'values')[0]
        selectionData = self._itemData[int(self.selectionID)]
        if selectionData[self._itemHeaders['On_Offer']]:
            offers = '£ %.2f'% (selectionData[self._itemHeaders['Price_Change']])
        else:
            offers = 'None'
        statsStr = str(selectionData[self._itemHeaders['Item_ID']]) + "\n\n" + \
                   selectionData[self._itemHeaders['Item_Name']]+"\n\n" + \
                   str(selectionData[self._itemHeaders['Date']])+"\n\n" + \
                   selectionData[self._itemHeaders['Category']]+"\n\n" + \
                   str(selectionData[self._itemHeaders['Batch_Stock']]) + "\n\n" + \
                   str(selectionData[self._itemHeaders['Original_Stock']])+"\n\n" + \
                   str(self.toBool(selectionData[self._itemHeaders['On_Offer']]))+"\n\n" + \
                   offers+"\n\n" + \
                   "£"+str(selectionData[self._itemHeaders['Price']])
        self.statsField.config(text=statsStr)

    def toBool(self, value):
        #function converts input to boolean
        if value:
            return True
        else:
            return False

    def sellItem(self, paypal=False):
        #function that validates inputs and adds the sale to the database
        try:
            int(self.Quantity.get())
            if not self.itemTree.selection() or not 0 < len(str(self.Quantity.get())) < 11 or not self.Quantity.get() > 0:
                int("a")
        except ValueError:
            tk.messagebox.showinfo("Select item", "Please select an item and enter a quantity (less than 11 "
                                                  "digits).")
            return

        self.selectionID = self.itemTree.item(self.itemTree.selection(), 'values')[0]
        selectionStock = self.itemTree.item(self.itemTree.selection(), 'values')[3]
        selectionData = self._itemData[int(self.selectionID)]
        if self.Quantity.get() > int(selectionStock):
            tk.messagebox.showerror("Not enough stock", "There is not enough stock of that product for that order "
                                                        "quantity.")
            return
        onOffer = selectionData[self._itemHeaders['On_Offer']]
        timelimit = selectionData[self._itemHeaders['Time_Limit']]
        price = selectionData[self._itemHeaders['Price_Selling']]
        quantity = self.Quantity.get()

        if onOffer and timelimit:
            #checks if there is an offer to charge the user at the offer price
            if dt.datetime.now() < timelimit:
                if selectionData[self._itemHeaders['Price_Change']]:
                    price = selectionData[self._itemHeaders['Price_Change']]
            else:
                if not tk.messagebox.askyesno("Offer expired", "The offer expired, do you wish to continue with the "
                                                               "purchase?"):
                    return


        if paypal:
            #if paypal button clicked, payment is made through paypal
            #paypal function called
            self.paymentSuccess = self.paypal(price, quantity, selectionData)
            if not self.paymentSuccess:
                tk.messagebox.showerror("Payment Failed", "Payment did not go through.")
                return
            uniqueID = self.paymentSuccess['salesID']
            self._transactionID = self.paymentSuccess['transaction']
            self._feeType = self.paymentSuccess['fee']
        else:
            uniqueID = self._controller.getID()
            self._transactionID = None
            self._feeType = None
        try:
            self.conn = MySQLConnection(**self.db)
            if self.conn.is_connected() != True:
                self._controller.log_event(self._controller.lineno(),
                                           "Database not connected contact the network admin.")
            self.c = self.conn.cursor()


            if timelimit:
                #if time limit of offer reached, toggle offer off
                if dt.datetime.now() >= timelimit:
                    self.c.execute("UPDATE offers SET On_Offer=False WHERE Item_ID=%(id)s", {'id':self.selectionID})

            #algorithm to iterate through each batch of the item and remove some of the stock until the purchased
            # quantity is reached
            batchStock = self._itemData[int(self.selectionID)][self._itemHeaders['Batch_Stock']]

            uniqueID += '.'
            uniquePart = 1
            remainder = self.Quantity.get()
            self.c.execute("SELECT Batch_ID, Original_Stock FROM all_inventory WHERE Item_ID=%(id)s AND "
                           "Date_Removed IS NULL ORDER BY Date", {'id':self.selectionID})
            Queue = self.c.fetchall()
            Queue = [list(elem) for elem in Queue]
            Queue[0][1] = batchStock
            for QueueCount, batch in enumerate(Queue):
                if remainder >= batch[1]:
                    self.c.execute("INSERT INTO sales VALUES (%(id)s, CURRENT_TIMESTAMP, %(item)s, %(price)s, "
                                   "%(quantity)s, %(income)s, %(staff)s,0, %(paypal)s, %(fee)s)",
                                   {'id': str(uniqueID+str(uniquePart)), 'item': batch[0],
                                    'price': price,
                                    'quantity': batch[1],
                                    'income': round(batch[1] * price, 2), 'staff': self.userID,
                                    'paypal': self._transactionID, 'fee':self._feeType})

                    self.c.execute("UPDATE all_inventory SET Date_Removed=CURRENT_TIMESTAMP WHERE Batch_ID=%(id)s",
                                   {'id': batch[0]})
                    if len(Queue) > 1:
                        self.c.execute("UPDATE all_inventory SET Queue=FALSE WHERE Batch_ID=%(id)s",
                                       {'id': Queue[QueueCount+1][0]})
                    remainder -= batch[1]
                    uniquePart += 1
                    if remainder == 0 and QueueCount+1 == len(Queue):
                        self.c.execute("DELETE FROM current_inventory WHERE Item_ID=%(id)s", {'id':self.selectionID})
                    else:
                        continue
                elif remainder > 0:
                    self.c.execute("INSERT INTO sales VALUES (%(id)s, CURRENT_TIMESTAMP, %(item)s, %(price)s, "
                                   "%(quantity)s, %(income)s, %(staff)s,0, %(paypal)s, %(fee)s)",
                                   {'id': str(uniqueID+str(uniquePart)), 'item': batch[0],
                                    'price': price,
                                    'quantity': remainder, 'income':
                                    round(remainder*price, 2),
                                    'staff': self.userID,
                                    'paypal': self._transactionID, 'fee':self._feeType})
                if QueueCount+1 == len(Queue):
                    totalStock = 0
                else:
                    totalStock = sum([elem[1] for elem in Queue[(QueueCount+1):]])
                #sets the new remaining stock quantity
                self.c.execute("UPDATE current_inventory SET Batch_Stock=%(stock)s, Total_Stock=%(total)s, "
                               "Last_Refresh=CURRENT_TIMESTAMP, Queue_length=%(queue)s WHERE "
                               "Item_ID=%(id)s",
                               {'stock': batch[1]-remainder, 'total':totalStock,
                                'queue':(len(Queue)-QueueCount)-1,
                                'id': self.selectionID})

                break
            tk.messagebox.showinfo("Success", "Sale completed successfully.")

        except Error as e:
            self._controller.log_event(e, self._controller.lineno())

        finally:
            self.conn.commit()
            self.conn.close()
            #refreshes treeview
            self.updateItemList()

    def restockItem(self):
        #function creates new batch of item in queue with specified restock quantity
        self.selectionID = self.itemTree.item(self.itemTree.selection(), 'values')[0]
        self.selectionName = self.itemTree.item(self.itemTree.selection(), 'values')[1]

        try:
            #edit database
            self.conn = MySQLConnection(**self.db)
            if self.conn.is_connected() != True:
                self._controller.log_event(self._controller.lineno(),
                                           "Database not connected contact the network admin.")
            self.c = self.conn.cursor()
            self._uniqueID = self._controller.getID()
            pricePaid = self._itemData[int(self.selectionID)][self._itemHeaders['Price_Selling']]
            self.c.execute("INSERT INTO all_inventory VALUES (%(id)s, CURRENT_TIMESTAMP, %(itemID)s, %(stock)s, "
                           "%(price)s, NULL, TRUE )", {'id':self._uniqueID, 'itemID':self.selectionID,
                                                       'stock':self.Quantity.get(), 'price':pricePaid})
            self.c.execute("UPDATE current_inventory SET Total_Stock=Total_Stock+%(stock)s, "
                           "Queue_length=Queue_length+1 WHERE Item_ID=%(id)s",
                           {'stock':self.Quantity.get(), 'id':self.selectionID})
            self._transID = self._controller.getID()
            description = 'Buying products to stock the inventory. ID='+self.selectionID+', Name='+self.selectionName
            self.c.execute("INSERT INTO income_expenditure VALUES(%(id)s, %(name)s, %(description)s"
                           ", %(price)s, %(quantity)s, %(total)s, "
                           "'expenditure', CURRENT_TIMESTAMP, %(batchID)s)",
                           {'id':self._transID,'name':'Stocking: '+self.selectionName, 'description':description,
                           'price':pricePaid, 'quantity':self.Quantity.get(), 'total':round(
                               pricePaid*self.Quantity.get(), 2), 'batchID':self._uniqueID})
            tk.messagebox.showinfo("Success", "Restock completed successfully.")

        except Error as e:
            self._controller.log_event(e, self._controller.lineno())

        finally:
            self.conn.commit()
            self.conn.close()
            #refresh database
            self.updateItemList()

    def discardItem(self):
        #delete quantity of item without making a sale
        self.selectionID = self.itemTree.item(self.itemTree.selection(), 'values')[0]
        selectionStock = self.itemTree.item(self.itemTree.selection(), 'values')[3]
        if self.Quantity.get() > int(selectionStock):
            tk.messagebox.showerror("Not enough stock", "There is not enough stock of that product to discard that "
                                                        "much.")
            return

        try:
            self.conn = MySQLConnection(**self.db)
            if self.conn.is_connected() != True:
                self._controller.log_event(self._controller.lineno(),
                                           "Database not connected contact the network admin.")

            self.c = self.conn.cursor()
            #algorithm to iterate through each batch and delete stock until the required quantity is reached
            batchStock = self._itemData[int(self.selectionID)][self._itemHeaders['Batch_Stock']]

            self.c.execute("SELECT Batch_ID, Original_Stock FROM all_inventory WHERE Item_ID=%(id)s AND "
                           "Date_Removed IS NULL ORDER BY Date", {'id':self.selectionID})
            Queue = self.c.fetchall()
            Queue = [list(elem) for elem in list(Queue)]
            Queue[0][1] = batchStock
            remainder = self.Quantity.get()
            for QueueCount, batch in enumerate(Queue):
                if remainder >= batch[1]:
                    self.c.execute("UPDATE all_inventory SET Date_Removed=CURRENT_TIMESTAMP WHERE Batch_ID=%(id)s",
                                   {'id': batch[0]})
                    if len(Queue) > 1:
                        self.c.execute("UPDATE all_inventory SET Queue=FALSE WHERE Batch_ID=%(id)s",
                                       {'id': Queue[QueueCount+1][0]})
                    remainder -= batch[1]
                    self._uniqueID = int(self._uniqueID) + 1
                    if remainder == 0 and QueueCount+1 == len(Queue):
                        self.c.execute("DELETE FROM current_inventory WHERE Item_ID=%(id)s", {'id':self.selectionID})
                    else:
                        continue

                if QueueCount+1 == len(Queue):
                    totalStock = 0
                else:
                    totalStock = sum([elem[1] for elem in Queue[QueueCount+1]])
                self.c.execute("UPDATE current_inventory SET Batch_Stock=%(stock)s, Total_Stock=%(total)s, "
                               "Last_Refresh=CURRENT_TIMESTAMP, Queue_length=%(queue)s WHERE "
                               "Item_ID=%(id)s",
                               {'stock': batch[1]-remainder, 'total':totalStock,
                                'queue':(len(Queue)-QueueCount)-1,
                                'id': self.selectionID})

                break

            tk.messagebox.showinfo("Success", "Items discarded successfully.")

        except Error as e:
            self._controller.log_event(e, self._controller.lineno())

        finally:
            self.conn.commit()
            self.conn.close()
            #refresh treeview
            self.updateItemList()

    def paypal(self, price, quantity, selection):
        #function used to carry out paypal payments for items
        selectionData = selection

        uniqueID = self._controller.getID()
        fee = None
        total = round(float(price)*int(quantity) ,2)
        if tk.messagebox.askyesno("Add fee", "Do you want to add the 3.4% + 20p PayPal fee to the purchase price?"):
            #asks if the user wants to add PayPal's fee on top of the purchase price
            fee = True
            onlineTax = total * 0.034
            onlineTax += 0.2
            onlineTax = round(onlineTax, 2)
            sale_price = round(total + onlineTax, 2)
        else:
            onlineTax = 0
            sale_price = round(total + onlineTax, 2)

        #constructs payment
        payment = paypal.Payment({
            "intent": "sale",
            "redirect_urls": {
                "return_url": website,
                "cancel_url": website
            },
            "payer": {
                "payment_method": "paypal",
            },
            "transactions": [
                {
                    "amount": {
                        "total": str(sale_price),
                        "currency": "GBP",
                        "details": {
                            "subtotal": str(total),
                            "handling_fee": str(onlineTax),
                        }
                    },
                    "description": "Buy a product; invest in yourself!",
                    "invoice_number": uniqueID,
                    "soft_descriptor": self.company['Company_Name']+" - "+uniqueID,
                    "item_list": {
                        "items": [
                            {
                                "name": selectionData[self._itemHeaders['Item_Name']],
                                "quantity": quantity,
                                "price": str(price),
                                "sku": str(selectionData[self._itemHeaders['Batch_ID']]) + " / " + str(selectionData[
                                    self._itemHeaders['Item_ID']]),
                                "currency": "GBP"
                            }
                        ]
                    }
                }
            ]
        })

        #creates payment url to gateway
        if payment.create():
            self._controller.log_event("Payment[%s] created successfully" % (payment.id), self._controller.lineno())
            # Redirect the user to given approval url
            for link in payment.links:
                if link.rel == "approval_url":
                    # Convert to str to avoid google appengine unicode issue
                    # https://github.com/paypal/rest-api-sdk-python/pull/58
                    approval_url = str(link.href)
                    self._controller.log_event("Redirect for approval: %s" % (approval_url), self._controller.lineno())
                    webb.open(approval_url)
        else:
            self._controller.log_event("Error while creating payment:", self._controller.lineno())
            self._controller.log_event(payment.error, self._controller.lineno())
            return

        if tk.messagebox.askyesno("Authorise payment", "Your browser should have opened a paypal payment page. Have "
                                                       "you entered your details and authorised the payment?"):
            try:
                # Payment ID obtained when creating the payment (following redirect)
                payment = paypal.Payment.find(payment.id)
                payerid = payment.payer.payer_info.payer_id

                # Execute payment with the payer ID from the create payment call (following redirect)
                if payment.execute({"payer_id": payerid}):
                    self._controller.log_event("Payment[%s] execute successfully" % (payment.id))
                    transactionid = payment.transactions[0].related_resources[0].sale.id
                    return {'transaction':transactionid, 'salesID':uniqueID, 'fee':fee}
                else:
                    self._controller.log_event(payment.error)
            except AttributeError:
                return {}
        else:
            return {}

    def xtraFunctions(self):
        #validates item and quanity inputs and calls function corresponding to action selected
        try:
            self.Quantity.get()
            if not self.itemTree.selection() or not 0 < len(str(self.Quantity.get())) < 11 or not self.Quantity.get() > 0:
                int("a")
        except ValueError:
            tk.messagebox.showinfo("Select item", "Please select an item and enter a quantity (less than 11 "
                                                  "digits).")
            return

        self.currentFunc = self.XFunctions.get()
        if self.currentFunc == "Sell":
            self.sellItem()
        elif self.currentFunc == "Restock":
            self.restockItem()
        elif self.currentFunc == "Discard Item":
            self.discardItem()
        else:
            tk.messagebox.showerror("No option", "Please select an option from the drop-down list.")

    def createSideBar(self):
        #creates button widgets that make up the side bar to perform functions with inventory items
        self.sidebar = tk.LabelFrame(self, text="Menu", bg=self._bg, bd=3)
        self.sidebar.place(relx=0.01, rely=0.1, relheight=0.4, relwidth=0.2)
        self.createB = tk.Button(self.sidebar, text="Create Item", font=self.fonts[2], command=self.createItem)
        self.createB.place(relx=0.1, rely=0.05, relheight=0.1, relwidth=0.8)
        self.BestSellingB = tk.Button(self.sidebar, text="Best to Worst Selling", font=self.fonts[2],
                                      command=self.viewTopSelling)
        self.BestSellingB.place(relx=0.1, rely=0.2, relheight=0.1, relwidth=0.8)
        self.WorstSellingB = tk.Button(self.sidebar, text="Price Match", font=self.fonts[2], command=self.priceMatch)
        self.WorstSellingB.place(relx=0.1, rely=0.35, relheight=0.1, relwidth=0.8)
        self.PriceMatchB = tk.Button(self.sidebar, text="Create Offer", font=self.fonts[2], command=self.createOffer)
        self.PriceMatchB.place(relx=0.1, rely=0.5, relheight=0.1, relwidth=0.8)
        self.CreateOfferB = tk.Button(self.sidebar, text="Refresh", font=self.fonts[2], command=self.updateItemList)
        self.CreateOfferB.place(relx=0.1, rely=0.65, relheight=0.1, relwidth=0.8)
        self.ArchivesB = tk.Button(self.sidebar, text="Download Archived Stock", font=self.fonts[2],
                                   command=self.downloadDB)
        self.ArchivesB.place(relx=0.1, rely=0.8, relheight=0.1, relwidth=0.8)

        if self.user['Access_Level'] > 3:
            #if access level greater than 3, disable create, price match and download-to-spreadsheet buttons
            self.createB.config(state='disabled')
            self.PriceMatchB.config(state='disabled')
            self.ArchivesB.config(state='disabled')

        self.sidebar.place_forget()

    def updateItemList(self):
        #function updates treeview with latest data from database
        self.itemTree.delete(*self.itemTree.get_children())
        try:
            #get data from database
            self.conn = MySQLConnection(**self.db)
            if self.conn.is_connected() != True:
                self._controller.log_event(self._controller.lineno(),
                                           "Database not connected contact the network admin.")
            self.c = self.conn.cursor()
            self.c.execute("SELECT current_inventory.*, all_inventory.Batch_ID, all_inventory.Date, "
                           "inventory.Item_Name, inventory.Category, all_inventory.Original_Stock, " 
                           "all_inventory.Price_Paid, ebay_price_history.Price, offers.On_Offer, "
                           "offers.Price_Change, offers.Time_Limit FROM current_inventory "
                           "INNER JOIN "
                                "all_inventory ON current_inventory.Item_ID = all_inventory.Item_ID "
                           "INNER JOIN "
                                "ebay_price_history ON all_inventory.Item_ID = ebay_price_history.Item_ID "
                           "INNER JOIN "
                                "offers ON ebay_price_history.Item_ID = offers.Item_ID "
                           "INNER JOIN "
                                "inventory ON offers.Item_ID = inventory.Item_ID"
                           " WHERE all_inventory.Date_Removed IS NULL AND "
                           "all_inventory.Queue=False AND ebay_price_history.Current=True")

            self._itemHeaders = {elem[1][0]:elem[0] for elem in enumerate(self.c.description)}
            self._itemData = {elem[0]:elem[0:] for elem in self.c.fetchall()}
            #populating treeview
            for item in self._itemData:
                self.itemTree.insert("", "end", values=[self._itemData[item][self._itemHeaders['Item_ID']],
                                                        self._itemData[item][self._itemHeaders['Item_Name']],
                                                        "£"+str(self._itemData[item][self._itemHeaders['Price_Selling']]),
                                                        int(self._itemData[item][self._itemHeaders[
                                                            'Batch_Stock']])+int(self._itemData[item][
                                                            self._itemHeaders['Total_Stock']]),
                                                        self._itemData[item][self._itemHeaders['Last_Refresh']]])

        except Error as e:
            self._controller.log_event(e, self._controller.lineno())

        finally:
            self.conn.commit()
            self.conn.close()
            self.treeChild = self.itemTree.get_children('')

    def createItem(self):
        #create frame to input new item details
        self.createFrame = tk.Frame(self, bg=self._bg)
        self.createFrame.place(relx=0, rely=0, relheight=1, relwidth=1)

        self._title2 = tk.Label(self.createFrame, text="Add Item", bg=self._bg, font=self.fonts[0])
        self._title2.place(relx=0.45, rely=0.05, relheight=0.05, relwidth=0.2)
        # Labels
        self._idL = tk.Label(self.createFrame, text="Item ID (if exists)", bg=self._bg, font=self.fonts[1], anchor='e')
        self._idL.place(relx=0.28, rely=0.15, relheight=0.05, relwidth=0.2)

        self._nameL = tk.Label(self.createFrame, text="Item Name", bg=self._bg, font=self.fonts[1], anchor='e')
        self._nameL.place(relx=0.28, rely=0.25, relheight=0.05, relwidth=0.2)
        self._quantityL = tk.Label(self.createFrame, text="Stock Quantity", bg=self._bg, font=self.fonts[1], anchor='e')
        self._quantityL.place(relx=0.28, rely=0.35, relheight=0.05, relwidth=0.2)
        self._pricepaidL = tk.Label(self.createFrame, text="Price Paid (per unit)", bg=self._bg, font=self.fonts[1],
                                    anchor='e')
        self._pricepaidL.place(relx=0.28, rely=0.45, relheight=0.05, relwidth=0.2)
        self._pricesellL = tk.Label(self.createFrame, text="Price Selling (per unit)", bg=self._bg, font=self.fonts[
            1], anchor='e')
        self._pricesellL.place(relx=0.28, rely=0.55, relheight=0.05, relwidth=0.2)
        self._categoryL = tk.Label(self.createFrame, text="Category", bg=self._bg, font=self.fonts[1], anchor='e')
        self._categoryL.place(relx=0.28, rely=0.65, relheight=0.05, relwidth=0.2)
        self._priceMatchL = tk.Label(self.createFrame, text="Price Match", bg=self._bg, font=self.fonts[1], anchor='e')
        self._priceMatchL.place(relx=0.28, rely=0.75, relheight=0.05, relwidth=0.2)

        # entry widgets
        self._ID = tk.StringVar()
        self._idE = tk.Entry(self.createFrame, textvariable=self._ID, bg=self._bg)
        self._idE.place(relx=0.5, rely=0.15, relheight=0.05, relwidth=0.2)
        #button used to load details from previously created item with similar attributes to the new desired one
        self._loadB = tk.Button(self.createFrame, text="Load", bg=self._bg, font=self.fonts[2],
                                command=self.loadItemData)
        self._loadB.place(relx=0.7, rely=0.15, relheight=0.05, relwidth=0.05)

        self._Name = tk.StringVar()
        self._nameE = tk.Entry(self.createFrame, textvariable=self._Name, bg=self._bg)
        self._nameE.place(relx=0.5, rely=0.25, relheight=0.05, relwidth=0.2)

        self._Quantity = tk.StringVar()
        self._quantityE = tk.Entry(self.createFrame, textvariable=self._Quantity, bg=self._bg)
        self._quantityE.place(relx=0.5, rely=0.35, relheight=0.05, relwidth=0.2)

        self._Cost = tk.StringVar()
        self._pricepaidE = tk.Entry(self.createFrame, textvariable=self._Cost, bg=self._bg)
        self._pricepaidE.place(relx=0.5, rely=0.45, relheight=0.05, relwidth=0.2)

        self._Price = tk.StringVar()
        self._pricesellE = tk.Entry(self.createFrame, textvariable=self._Price, bg=self._bg)
        self._pricesellE.place(relx=0.5, rely=0.55, relheight=0.05, relwidth=0.2)

        self._Category = tk.StringVar()
        self._categoryE = ttk.Combobox(self.createFrame, textvariable=self._Category)
        self._categoryE.place(relx=0.5, rely=0.65, relheight=0.05, relwidth=0.2)
        self.getOptions()

        self._PriceMatch = tk.StringVar()
        self._priceMatchE = tk.Entry(self.createFrame, textvariable=self._PriceMatch, bg=self._bg, state='disabled')
        self._priceMatchE.place(relx=0.5, rely=0.75, relheight=0.05, relwidth=0.2)

        #return to the main inventory page
        self._cancelB = tk.Button(self.createFrame, text="Cancel", bg=self._bg, font=self.fonts[2],
                                  command=lambda: self.back([self.createFrame], False))
        self._cancelB.place(relx=0.43, rely=0.85, relheight=0.05, relwidth=0.1)
        #proceed to frame to input eBay search criteria
        self._nextB = tk.Button(self.createFrame, text="Next", bg=self._bg, font=self.fonts[2], command=self.validate1)
        self._nextB.place(relx=0.59, rely=0.85, relheight=0.05, relwidth=0.1)

        #after price comparison, this button becomes unlocked and used to save item details to database
        self._confirmB = tk.Button(self.createFrame, text="Save Item", state='disabled', bg=self._bg,
                                   font=self.fonts[2], command=lambda: self.saveItemData())
        self._confirmB.place(relx=0.85, rely=0.93, relheight=0.05, relwidth=0.1)

    def createItemPrice(self, itemName=None):
        #creates frame to select search criteria to find matching eBay price
        self.PriceFrame = tk.Frame(self.createFrame, bg=self._bg)
        self.PriceFrame.place(relx=0, rely=0, relheight=1, relwidth=1)
        self._title3 = tk.Label(self.PriceFrame, text="Price Match", bg=self._bg, font=self.fonts[0])
        self._title3.place(relx=0.4, rely=0.05, relheight=0.05, relwidth=0.2)

        #returns to item details frame
        self._backB1 = tk.Button(self.PriceFrame, text="Back", bg=self._bg, font=self.fonts[2],
                                command=lambda: self.back([self.PriceFrame], False))
        self._backB1.place(relx=0.05, rely=0.9, relheight=0.05, relwidth=0.1)

        #proceed to match the price and view resulting listings
        self._resultsB = tk.Button(self.PriceFrame, text="Match Price", bg=self._bg, font=self.fonts[2],
                                   command=self.matchPrice)
        self._resultsB.place(relx=0.85, rely=0.9, relheight=0.05, relwidth=0.1)


        self._priceL = tk.Label(self.PriceFrame, text="My Price: £" + self._Price.get(), bg=self._bg,
                                font=self.fonts[2])
        self._priceL.place(relx=0.4, rely=0.85, relheight=0.05, relwidth=0.2)

        # eBay Price Comparison Entry
        self.ebayFrame = tk.LabelFrame(self.PriceFrame, text="eBay Price Match", bg=self._bg, bd=2)
        self.ebayFrame.place(relx=0.3, rely=0.1, relheight=0.75, relwidth=0.4)
        self._ebSearchL = tk.Label(self.ebayFrame, text='Search', bg=self._bg, font=self.fonts[1])
        self._ebSearchL.place(relx=0.05, rely=0.05, relheight=0.05, relwidth=0.2)
        self._ebSearch = tk.StringVar()
        self._ebSearchE = tk.Entry(self.ebayFrame, textvariable=self._ebSearch, bg=self._bg)
        self._ebSearchE.place(relx=0.1, rely=0.1, relheight=0.08, relwidth=0.8)
        if itemName:
            self._ebSearch.set(itemName)

        self._ebFilterL = tk.Label(self.ebayFrame, text="Filters", bg=self._bg, font=self.fonts[1])
        self._ebFilterL.place(relx=0.05, rely=0.2, relheight=0.05, relwidth=0.2)

        self._ebConditionL = tk.Label(self.ebayFrame, text="Condition", bg=self._bg, font=self.fonts[2], anchor='e')
        self._ebConditionL.place(relx=0.05, rely=0.25, relheight=0.05, relwidth=0.25)
        self._ebCondition = tk.StringVar()
        self._ebCondition.set('All')
        self._ebConditionE = ttk.OptionMenu(self.ebayFrame, self._ebCondition, *('','All', 'New', 'New other (see '
                                                                                               'details)',
                                                                                'New with defects', 'Manufacturer '
                                                                                                    'refurbished',
                                                                                'Seller refurbished',
                                                                                'Used', 'Very Good', 'Good',
                                                                                'Acceptable',
                                                                                'For parts or not working'))
        self._ebConditionE.place(relx=0.35, rely=0.25, relheight=0.05, relwidth=0.3)

        self._ebListingL = tk.Label(self.ebayFrame, text="Listing Type", bg=self._bg,
                                    font=self.fonts[2], anchor='e')
        self._ebListingL.place(relx=0.05, rely=0.35, relheight=0.05, relwidth=0.25)
        self._ebListing = tk.StringVar()
        self._ebListing.set('All')
        self._ebListingE = ttk.OptionMenu(self.ebayFrame, self._ebListing, *('','All', 'FixedPrice', 'Auction',
                                                                            'AuctionWithBIN', 'Classified',
                                                                            'StoreInventory'))
        self._ebListingE.place(relx=0.35, rely=0.35, relheight=0.05, relwidth=0.3)

        self._ebMinPriceL = tk.Label(self.ebayFrame, text="Min. Price", bg=self._bg, font=self.fonts[2], anchor='e')
        self._ebMinPriceL.place(relx=0.05, rely=0.45, relheight=0.05, relwidth=0.25)
        self._ebMinPrice = tk.DoubleVar()
        self._ebMinPrice.set('0')
        self._ebMinPriceE = tk.Entry(self.ebayFrame, textvariable=self._ebMinPrice, bg=self._bg)
        self._ebMinPriceE.place(relx=0.35, rely=0.45, relheight=0.05, relwidth=0.1)

        self._ebMaxPriceL = tk.Label(self.ebayFrame, text="Max. Price", bg=self._bg, font=self.fonts[2], anchor='e')
        self._ebMaxPriceL.place(relx=0.05, rely=0.55, relheight=0.05, relwidth=0.25)
        self._ebMaxPrice = tk.DoubleVar()
        self._ebMaxPrice.set(-1)
        self._ebMaxPriceE = tk.Entry(self.ebayFrame, textvariable=self._ebMaxPrice, bg=self._bg)
        self._ebMaxPriceE.place(relx=0.35, rely=0.55, relheight=0.05, relwidth=0.1)

        self._ebReturnsL = tk.Label(self.ebayFrame, text="Returns Accepted", bg=self._bg, font=self.fonts[2],
                                    anchor='e')
        self._ebReturnsL.place(relx=0.05, rely=0.65, relheight=0.05, relwidth=0.25)
        self._ebReturns = tk.StringVar()
        self._ebReturns.set('All')
        self._ebReturnsE = ttk.OptionMenu(self.ebayFrame, self._ebReturns, *('','All', 'True', 'False'))
        self._ebReturnsE.place(relx=0.35, rely=0.65, relheight=0.05, relwidth=0.3)

        self._ebTopSellerL = tk.Label(self.ebayFrame, text="Top Rated Seller", bg=self._bg, font=self.fonts[2],
                                      anchor='e')
        self._ebTopSellerL.place(relx=0.05, rely=0.75, relheight=0.05, relwidth=0.25)
        self._ebTopSeller = tk.StringVar()
        self._ebTopSeller.set('All')
        self._ebTopSellerE = ttk.OptionMenu(self.ebayFrame, self._ebTopSeller, *('','All', 'True', 'False'))
        self._ebTopSellerE.place(relx=0.35, rely=0.75, relheight=0.05, relwidth=0.3)

        self._ebStoreTypeL = tk.Label(self.ebayFrame, text="Business Type", bg=self._bg, font=self.fonts[2], anchor='e')
        self._ebStoreTypeL.place(relx=0.05, rely=0.85, relheight=0.05, relwidth=0.25)
        self._ebStoreType = tk.StringVar()
        self._ebStoreType.set('All')
        self._ebStoreTypeE = ttk.OptionMenu(self.ebayFrame, self._ebStoreType, *('','All', 'Business', 'Private'))
        self._ebStoreTypeE.place(relx=0.35, rely=0.85, relheight=0.05, relwidth=0.3)

    def matchPrice(self):
        #validate search parameter inputs and display resulting price match records in treeview table
        if not 0 < len(self._ebSearch.get()) <= 255:
            tk.messagebox.showerror("Invalid Input", "Enter a search term in the Search box that is no more than 255 "
                                            "characters.")
            return

        try:
            float(self._ebMinPrice.get())
            float(self._ebMaxPrice.get())
        except ValueError:
            tk.messagebox.showerror("Ensure there are no letters in the minimum and maximum price field. Only enter "
                                    "integers and decimals.")
            return
        if float(self._ebMinPrice.get()) < 0 or float(self._ebMaxPrice.get()) < -1:
            tk.messagebox.showerror("Invalid input","With the exception that -1 for maximum price indicates no maximum "
                                              "price, neither inputs should have a negative value.")

            return


        self.MatchFrame = tk.Frame(self.PriceFrame, bg=self._bg)
        self.MatchFrame.place(relx=0, rely=0, relheight=1, relwidth=1)

        self._instruct1 = tk.Label(self.MatchFrame, text="Choose the eBay listing that is/represents your item the closest. "
                                              "Double click to view the webpage to get more product details. ",
                                   bg=self._bg,
                                   font=self.fonts[2],
                                   wraplength=350,
                                   justify='center')
        self._instruct1.place(relx=0.3, rely=0.05, relheight=0.08, relwidth=0.4)


        self._treecolumns = ['Name', 'Category', 'Price', 'Condition', 'Listing Type', 'Listing Duration (days)',
                             'Seller Feedback', 'Seller Score', 'Top Seller', 'URL']
        #displays listings with similar attributes as item
        self.resultTree = ttk.Treeview(self.MatchFrame, height=20, columns=self._treecolumns, show='headings')
        for column in self._treecolumns:
            self.resultTree.heading(column, text=column, command=lambda col=column: self._controller.sortItem(
                col,self.resultTree))
            self.resultTree.column(column, stretch='yes', width=8)
        self.resultTree.place(relx=0.05, rely=0.15, relheight=0.65, relwidth=0.9)
        self.resultTree.bind("<Double-1>", lambda e:self.viewItemURL())
        self.resultTree.bind("<Return>", lambda e:self.chooseItem(self.resultTree.selection()))

        self.resultScroll = ttk.Scrollbar(self.MatchFrame, orient="vertical", command=self.resultTree.yview)
        self.resultScroll.place(relx=0.95, rely=0.15, relheight=0.65)

        self.resultTree.configure(yscrollcommand=self.resultScroll.set)
        #returns back to frame to input search parameters
        self._backB3 = tk.Button(self.MatchFrame, text="Back", bg=self._bg, font=self.fonts[2],
                                command=lambda: self.back([self.MatchFrame], False))
        self._backB3.place(relx=0.05, rely=0.9, relheight=0.05, relwidth=0.1)
        #select a listing and save it as the eBay price
        self._chooseB = tk.Button(self.MatchFrame, text="Choose", bg=self._bg,font=self.fonts[2],
                                   command=lambda:self.chooseItem(self.resultTree.selection()))
        self._chooseB.place(relx=0.85, rely=0.9, relheight=0.05, relwidth=0.1)

        self.parameters = {'min price':self._ebMinPrice.get(), 'max price':self._ebMaxPrice.get(),
                           'condition':self._ebCondition.get(), 'listing':self._ebListing.get(),
                           'returns':self._ebReturns.get(), 'seller rating':self._ebTopSeller.get(),
                           'business type':self._ebStoreType.get(), 'search':self._ebSearch.get()}

        #call getEBayData to match price with eBay api then iterate through results and add to treeview
        self.items = self.getEBayData(**self.parameters)
        for index, item in enumerate(self.items):
            self.listingDur = dt.datetime.strptime(item.endtime.string.lower(),
                                                   '%Y-%m-%dt%H:%M:%S.%fz') - dt.datetime.strptime(
                item.starttime.string.lower(), '%Y-%m-%dt%H:%M:%S.%fz')
            self.listingDur = self.listingDur.days
            self.url = item.viewitemurl.string.lower()
            self._itemURLS["EBlink" + str(index)] = self.url
            self.resultTree.insert('', 'end', text='ebay', values=(item.title.string.lower(),
                                                                   item.categoryname.string.lower(),
                                                                   float(item.currentprice.string),
                                                                   item.conditiondisplayname.string.lower(),
                                                                   item.listingtype.string.lower(),
                                                                   self.listingDur, item.feedbackscore.string.lower(),
                                                                   item.positivefeedbackpercent.string.lower(),
                                                                   item.topratedseller.string.lower(),
                                                                   "EBlink" + str(index)))

    def getEBayData(self, **params):
        #function to use passed params to get matching ebay listings
        eb_api = eBayConn(appid=self.ebayID['app_id'], siteid='EBAY-GB', config_file=None)
        self._itemURLS = {}

        api_filters = [{'name': 'country', 'value': 'UK'}]
        for filter in [['MinPrice', params['min price']], ['MaxPrice', params['max price']]]:
            if int(filter[1]) > 0:
                api_filters.append({'name': filter[0], 'value': filter[1]})

        for filter in [['Condition', params['condition']], ['ListingType', params['listing']],
                       ['ReturnsAcceptedOnly', params['returns']], ['TopRatedSellerOnly', params[
                'seller rating']], ['SellerBusinessType', params['business type']]]:
            if filter[1] != 'All':
                api_filters.append({'name': filter[0], 'value': filter[1]})
        #constructing request
        api_request = {
            'keywords': params['search'],
            'outputSelector': 'SellerInfo',
            'itemFilter': api_filters,
            'paginationInput': {
                'entriesPerPage': '10',
                'pageNumber': '1'
            }
        }
        #making ebay api request
        api_response = eb_api.execute('findItemsByKeywords', api_request)
        #parse the response
        soup = BeautifulSoup(api_response.content, 'lxml')
        return soup.find_all('item')

    def chooseItem(self, item):
        #function that selects an eBay listing for price match
        if not item:
            tk.messagebox.showerror("Select item", "Please select a product from the table that resembles yours the "
                                                  "closest.")
            return
        self.selection = self.resultTree.item(item, 'values')
        #sets the price match value in the match item details edit page
        self._priceMatchE.config(state='normal')
        self._PriceMatch.set(self.selection[2])
        self._priceMatchE.config(state='disabled')
        #allows user to click the 'Save Item' button
        self._confirmB.config(state='normal')
        self.back([self.PriceFrame, self.MatchFrame])

    def viewItemURL(self):
        #function that opens ebay item listing url when double clicked in the price match treeview
        if not self.resultTree.selection():
            tk.messagebox.showerror("Select item", "Please select a listing you wish to see.")
            return

        self._selectedItem = self.resultTree.item(self.resultTree.selection(), 'values')
        webb.open(self._itemURLS[self._selectedItem[-1]])

    def saveItemData(self):
        #function to save item details to database
        if not self.validate1(True):
            return
        try:
            self.conn = MySQLConnection(**self.db)
            if self.conn.is_connected() != True:
                self._controller.log_event(self._controller.lineno(),
                                           "Database not connected contact the network admin.")
            self.c = self.conn.cursor()

            self.c.execute("SELECT * FROM inventory WHERE (Item_Name=%(name)s OR "
                           "Item_ID=%(id)s)", {'name':self._Name.get(), 'id':self._ID.get()})
            self._currentItem = self.c.fetchall()
            self._itemID = False
            if self._currentItem:
                for version in self._currentItem:
                    #if ID or name is same as already saved item in database, the program gives the option to load
                    # previously stored price match history
                    if tk.messagebox.askyesno("Product exists", "Do you want to load price match data from this "
                                                                "previous item: \nID: "+str(version[0])+"\nName: "
                                                                 +version[1]+"\nCategory: "+version[2]):
                        self._itemID = version[0]
                        break

            #stores data in database
            if not self._itemID:
                self.c.execute("INSERT INTO current_inventory VALUES (NULL, %(stock)s, %(total)s, %(price)s, "
                               "CURRENT_TIMESTAMP, 0)",
                               {'stock':self._Quantity.get(), 'total':0, 'price':self._Price.get()})
                self._itemID = self.c.lastrowid
                isQueue = False
            else:
                self.c.execute("SELECT Original_Stock FROM all_inventory WHERE Date_Removed IS NULL AND Item_ID=%("
                               "id)s ORDER BY Date", {'id':self._itemID})
                self._queue = [elem[0] for elem in self.c.fetchall()]
                if self._queue:
                    #if there is only 1 item in queue then sum(self._queue[1:]) will not work
                    self._totalStock = sum(self._queue) - self._queue[0]
                    self._totalStock += self._Quantity.get()
                    isQueue = True
                    self.c.execute("UPDATE current_inventory SET Total_Stock=%(total)s, "
                                   "Price_Selling=%(price)s, Last_Refresh=CURRENT_TIMESTAMP, "
                                   "Queue_length=%(queue)s WHERE Item_ID=%(id)s",
                                                                        {'id': self._itemID,
                                                                        'total': self._totalStock,
                                                                        'queue': len(self._queue),
                                                                         'price':self._Price.get()})

                else:
                    self._totalStock = 0
                    isQueue = False
                self.c.execute("INSERT INTO current_inventory VALUES (%(id)s, %(stock)s, 0, %(price)s, "
                               "CURRENT_TIMESTAMP, 0)", {'id':self._itemID, 'stock':self._Quantity.get(),
                                                                 'price':self._Price.get()})

            self._uniqueID = self._controller.getID()

            self.c.execute("INSERT INTO all_inventory VALUES (%(uniqID)s, CURRENT_TIMESTAMP, %(id)s,"
                           "%(quantity)s, %(cost)s, NULL, %(queue)s)",
                           {'uniqID':self._uniqueID, 'id': self._itemID,
                            'quantity':self._Quantity.get(), 'cost':self._Cost.get(), 'queue':isQueue})

            self.c.execute("INSERT INTO inventory (Item_ID, Item_Name, Category) VALUES (%(id)s, %(name)s, "
                           "%(cat)s) ON DUPLICATE KEY UPDATE Item_Name=%(name)s, Category=%(cat)s",
                           {'id':self._itemID, 'name':self._Name.get(),'cat':self._Category.get(),})

            self.c.execute("INSERT INTO offers (Item_ID, On_Offer, Price_Change, Time_Limit) VALUES (%(id)s, False, "
                           "NULL, NULL) ON DUPLICATE KEY UPDATE On_Offer=False", {'id':self._itemID})


            self.c.execute("UPDATE ebay_price_history SET Current=False WHERE Item_ID=%(id)s AND Current=True",
                           {'id':self._itemID})

            self.c.execute("INSERT INTO ebay_price_history VALUES (%(PriceID)s, CURRENT_TIMESTAMP, %(ItemID)s, "
                           "%(ebay)s, 'manual', True)",
                           {'PriceID': self._uniqueID, 'ItemID':self._itemID, 'ebay':self._PriceMatch.get()})

            self.c.execute("INSERT INTO ebay_search_data (Item_ID, Search_Term, Item_Condition, Listing, AllowReturns, "
                           "Seller_Rating, Business_Type,Min_Price, Max_Price, Last_Updated) VALUES (%(id)s, "
                           "%(search)s, %(condition)s, %(listing)s, %(returns)s, %(rating)s, %(business)s, %(min)s, "
                           "%(max)s, CURRENT_TIMESTAMP) ON DUPLICATE KEY UPDATE Search_Term=%(search)s, "
                           "Item_Condition=%(condition)s, Listing=%(listing)s, AllowReturns=%(returns)s, "
                           "Seller_Rating=%(rating)s, Business_Type=%(business)s, Min_Price=%(min)s, Max_Price=%(max)s, "
                           "Last_Updated=CURRENT_TIMESTAMP",
                           {'id':self._itemID, 'search':self.parameters['search'], 'condition':self.parameters['condition'],
                            'listing':self.parameters['listing'], 'returns':self.parameters['returns'],
                            'rating':self.parameters['seller rating'], 'business':self.parameters['business type'],
                            'min':self.parameters['min price'], 'max':self.parameters['max price']})

            description = 'Buying products to stock the inventory. ID='+str(self._itemID)+', Name='+self._Name.get()

            self.c.execute("INSERT INTO income_expenditure VALUES(%(id)s, %(name)s, %(description)s"
                           ", %(price)s, %(quantity)s, %(total)s, "
                           "'expenditure', CURRENT_TIMESTAMP, %(batchID)s)",
                           {'id':self._uniqueID,'name':'Stocking: '+self._Name.get(), 'description':description,
                           'price':self._Price.get(), 'quantity':self._Quantity.get(), 'total':float(
                self._Price.get())*int(self._Quantity.get()), 'batchID':self._uniqueID})

            tk.messagebox.showinfo("Success", "Item successfully added.")



        except Error as e:
            self._controller.log_event(e, self._controller.lineno())

        finally:
            self.conn.commit()
            self.conn.close()
            self.back([self.createFrame])

    def getOptions(self):
        #functions gets options for product category for when adding a new item
        try:
            self.conn = MySQLConnection(**self.db)
            if self.conn.is_connected() != True:
                self._controller.log_event("Database not connected contact the network admin.",
                                           self._controller.lineno())
            self.c = self.conn.cursor()
            #get all different categories in database
            self.c.execute("SELECT DISTINCT Category FROM inventory")
            self._catOptions = [elem[0] for elem in self.c.fetchall()]
            self._categoryE.config(values=self._catOptions)

        except Error as e:
            self._controller.log_event(e, self._controller.lineno())
        finally:
            self.conn.commit()
            self.conn.close()

    def loadItemData(self):
        #function to populate item creation fields with data from already created item
        if len(self._ID.get()) <= 10:
            #validates if inputted item ID exists
            try:
                self.conn = MySQLConnection(**self.db)
                if self.conn.is_connected() != True:
                    self._controller.log_event(self._controller.lineno(),
                                               "Database not connected contact the network admin.")
                self.c = self.conn.cursor()



                self.c.execute("SELECT inventory.Item_Name, all_inventory.Original_Stock, all_inventory.Price_Paid, "
                               "inventory.Category, ebay_price_history.Price FROM all_inventory "
                               "INNER JOIN "
                               "inventory ON all_inventory.Item_ID = inventory.Item_ID "
                               "INNER JOIN "
                               "ebay_price_history ON inventory.Item_ID = ebay_price_history.Item_ID "
                               "WHERE all_inventory.Item_ID=%(id)s AND ebay_price_history.CURRENT=TRUE ORDER BY all_inventory.Date DESC LIMIT 1", {'id': self._ID.get()})
                self.item_data = self.c.fetchall()
                #populates fields
                if self.item_data:
                    self.item_data = self.item_data[0]
                    self._Name.set(self.item_data[0])
                    self._Quantity.set(self.item_data[1])
                    self._Cost.set(self.item_data[2])
                    self._Category.set(self.item_data[3])
                    self._priceMatchE.config(state='normal')
                    self._PriceMatch.set(self.item_data[4])
                    self._priceMatchE.config(state='disabled')
                    self.c.execute("SELECT Price_Selling FROM current_inventory "
                                   "WHERE Item_ID=%(id)s LIMIT 1", {'id': self._ID.get()})
                    itemPrice = self.c.fetchall()
                    if itemPrice:
                        self._Price.set(itemPrice[0][0])

                else:
                    tk.messagebox.showerror("Incorrect Input", "Item ID doesn't exist")

            except Error as e:
                self._controller.log_event(e, self._controller.lineno())

            finally:
                self.conn.commit()
                self.conn.close()
        else:
            tk.messagebox.showerror("Incorrect Input", "Item ID is not valid")

    def validate1(self, justCheck=False):
        #validates the inputs from creating a new item
        if not 0 < len(self._Name.get()) <= 255:
            tk.messagebox.showerror("Invalid Input", "Ensure the name is no greater than 255 characters.")
            return
        if not self._Quantity.get().isnumeric():
            tk.messagebox.showerror("Invalid Input", "Ensure the quantity is an integer.")
            return
        elif not 0 < len(self._Quantity.get()) < 5:
            tk.messagebox.showerror("Invalid Input", "Ensure the quantity is non-zero and no greater than 5 "
                                                     "digits.")
            return

        try:
            float(self._Cost.get())
            float(self._Price.get())
        except ValueError:
            tk.messagebox.showerror("Invalid Input", "Ensure the cost and price is a decimal number only (not £ "
                                                     "sign).")
            return

        if not 0 < len(str(self._Cost.get())) <= 12:
            tk.messagebox.showerror("Invalid Input", "Ensure the cost is non-zero and no greater than £9 billion.")
            return

        money = self._Cost.get().split(".")
        if len(money) == 2:
            if not 0 <= len(money[1]) <= 2:
                tk.messagebox.showerror("Invalid Input", "Ensure the cost is to 2 decimal places only.")
                return

        if not 0 < len(str(self._Price.get())) <= 12:
            tk.messagebox.showerror("Invalid Input", "Ensure the price you're selling at is non-zero and no greater "
                                                     "than £9 billion.")
            return

        money = self._Price.get().split(".")
        if len(money) == 2:
            if not 0 <= len(money[1]) <= 2:
                tk.messagebox.showerror("Invalid Input", "Ensure the price is to 2 decimal places only.")
                return


        if not 0 < len(self._Category.get()) <= 255:
            tk.messagebox.showerror("Invalid Input", "Ensure category is no greater than 255 characters.")
            return

        if not justCheck:
            #calls function to navigate to the page to enter search terms for eBay price match
            self.createItemPrice(itemName=self._Name.get())
        return True

    def openSideBar(self):
        #toggles the side bar from display
        if self.sidebar.winfo_ismapped():
            self.sidebar.place_forget()
        else:
            self.sidebar.place(relx=0.01, rely=0.1, relheight=0.4, relwidth=0.2)
            self.sidebar.focus_set()

    def createOffer(self):
        #creates frame and widgets to input offer data
        if not self.itemTree.selection():
            tk.messagebox.showerror("No selection", "Select an item to create an offer for it.")
            return
        self.selectionID = int(self.itemTree.item(self.itemTree.selection(), 'values')[0])
        stock = self._itemData[self.selectionID][self._itemHeaders['Batch_Stock']]+self._itemData[self.selectionID][
            self._itemHeaders['Total_Stock']]
        price = self._itemData[self.selectionID][self._itemHeaders['Price_Selling']]
        self.offerFrame = tk.Frame(self, bg=self._bg)
        self.offerFrame.place(relx=0, rely=0, relheight=1, relwidth=1)
        self._title4 = tk.Label(self.offerFrame, text="Create Offer", bg=self._bg, font=self.fonts[0])
        self._title4.place(relx=0.4, rely=0.02, relheight=0.05, relwidth=0.2)

        #checkbutton toggles offer activation
        self._OnOffer = tk.IntVar()
        self._onOfferB = tk.Checkbutton(self.offerFrame, text="Activate Offer", bg='red', selectcolor='green', bd=0, highlightthickness=0,
                                          indicatoron=False, font=self.fonts[4], variable=self._OnOffer)
        self._onOfferB.place(relx=0.43, rely=0.11, relheight=0.05, relwidth=0.14)


        self._PriceChange = tk.DoubleVar()
        self._priceChangeL = tk.Label(self.offerFrame, text="Price Change", bg=self._bg, font=self.fonts[1])
        self._priceChangeL.place(relx=0.3, rely=0.2, relheight=0.05, relwidth=0.1)
        self._priceChangeE = tk.Entry(self.offerFrame, textvariable=self._PriceChange, bg=self._bg,
                                        font=self.fonts[2])
        self._priceChangeE.place(relx=0.3, rely=0.26,relheight=0.05, relwidth=0.1)
        self._PriceChange.set(price)

        self._itemStatsL = tk.Label(self.offerFrame, text="Current Price: "+str(price)+"\nCurrent Stock: "+str(stock),
                                bg=self._bg,
                                fg='gray', font=self.fonts[2])
        self._itemStatsL.place(relx=0.29, rely=0.31, relheight=0.05, relwidth=0.12)


        self._TimeLimit = [tk.IntVar(), tk.IntVar(), tk.IntVar(), tk.IntVar(), tk.IntVar()]
        self._timeL = tk.Label(self.offerFrame, text="Time Limit", bg=self._bg, font=self.fonts[1])
        self._timeL.place(relx=0.6, rely=0.2, relheight=0.05, relwidth=0.1)
        self._timeE1 = tk.Entry(self.offerFrame, textvariable=self._TimeLimit[0], font=self.fonts[2], bg=self._bg)
        self._timeE1.place(relx=0.5, rely=0.26, relheight=0.05, relwidth=0.05)
        tk.Label(self.offerFrame, text="/", font=self.fonts[2], bg=self._bg).place(relx=0.55, rely=0.26,
                                                                                   relheight=0.05, relwidth=0.01)
        self._timeE2 = tk.Entry(self.offerFrame, textvariable=self._TimeLimit[1], font=self.fonts[2], bg=self._bg)
        self._timeE2.place(relx=0.56, rely=0.26, relheight=0.05, relwidth=0.05)
        tk.Label(self.offerFrame, text="/", font=self.fonts[2], bg=self._bg).place(relx=0.61, rely=0.26,
                                                                                   relheight=0.05, relwidth=0.01)

        self._timeE3 = tk.Entry(self.offerFrame, textvariable=self._TimeLimit[2], font=self.fonts[2], bg=self._bg)
        self._timeE3.place(relx=0.62, rely=0.26, relheight=0.05, relwidth=0.05)

        self._timeE4 = tk.Entry(self.offerFrame, textvariable=self._TimeLimit[3], font=self.fonts[2], bg=self._bg)
        self._timeE4.place(relx=0.7, rely=0.26, relheight=0.05, relwidth=0.05)
        tk.Label(self.offerFrame, text=":", font=self.fonts[2], bg=self._bg).place(relx=0.75, rely=0.26,
                                                                                   relheight=0.05, relwidth=0.01)

        self._timeE5 = tk.Entry(self.offerFrame, textvariable=self._TimeLimit[4], font=self.fonts[2], bg=self._bg)
        self._timeE5.place(relx=0.76, rely=0.26, relheight=0.05, relwidth=0.05)

        self._priceL = tk.Label(self.offerFrame, text="DD/MM/YYYY HH:MM\nUse all 0s for no time limit", bg=self._bg,
                                fg='gray', font=self.fonts[2])
        self._priceL.place(relx=0.54, rely=0.31, relheight=0.07, relwidth=0.2)

        #save offer to database
        self._saveB = tk.Button(self.offerFrame, text="Save", bg=self._bg, font=self.fonts[2], command=self.saveOffer)
        self._saveB.place(relx=0.52, rely=0.45, relheight=0.05, relwidth=0.08)
        #returns to main inventory page
        self._backB2 = tk.Button(self.offerFrame, text="Back", bg=self._bg, font=self.fonts[2], command=lambda: self.back(frame=[self.offerFrame]))
        self._backB2.place(relx=0.43, rely=0.45, relheight=0.05, relwidth=0.08)

        currentOffer = self.fillOffer(self.selectionID)
        if currentOffer:
            #populates fields with existing data
            self._OnOffer.set(currentOffer['offer'])
            self._PriceChange.set(currentOffer['price'])
            if currentOffer['time']:
                self._TimeLimit[0].set(currentOffer['time'].day)
                self._TimeLimit[1].set(currentOffer['time'].month)
                self._TimeLimit[2].set(currentOffer['time'].year)
                self._TimeLimit[3].set(currentOffer['time'].hour)
                self._TimeLimit[4].set(currentOffer['time'].minute)

    def saveOffer(self):
        #validate inputs and save new offer to database
        success = False
        try:
            self.conn = MySQLConnection(**self.db)
            if self.conn.is_connected() != True:
                self._controller.log_event(self._controller.lineno(),
                                           "Database not connected contact the network admin.")
            self.c = self.conn.cursor()

            try:
                time = False
                for value in self._TimeLimit:
                    if not value.get() == 0:
                        time = True
                        break
                if time:
                    dateFull = dt.datetime(day=self._TimeLimit[0].get(), month=self._TimeLimit[1].get(), year=self._TimeLimit[2].get(),
                                           hour=self._TimeLimit[3].get(), minute=self._TimeLimit[4].get())
                    if dateFull > dt.datetime.today()+dt.timedelta(days=93) or dateFull < dt.datetime.today():
                        tk.messagebox.showerror("Invalid Input", "The offer cannot have a time limit of greater than "
                                                                 "93 days (~3 months) and cannot be set to before the current date and time")
                        return
                else:
                    dateFull = None
            except (TypeError, ValueError):
                tk.messagebox.showerror("Invalid input", "Ensure you either enter a valid date under the time limit, "
                                                         "or all zeroes for no time limit.")
                return
            try:
                float(self._PriceChange.get())
                number, decimal = str(self._PriceChange.get()).split(".")
                if len(decimal) > 2:
                    raise ValueError
            except ValueError:
                tk.messagebox.showerror("Invalid input", "Ensure the price change is a number with 2 decimal places "
                                                         "maximum.")
                return
            #update database
            self.c.execute("UPDATE offers SET On_Offer=%(on)s, Price_Change=%(price)s, Time_Limit=%("
                           "time)s WHERE Item_ID=%(id)s", {'id':self.selectionID, 'on':self._OnOffer.get(),
                                      'price':self._PriceChange.get(), 'time':dateFull})

            tk.messagebox.showinfo("Success", "Offer created!")
            success = True
        except Error as e:
            self._controller.log_event(e, self._controller.lineno())
        finally:
            self.conn.commit()
            self.conn.close()
            if success:
                #if save successful, exit from offer frame and update treeview
                self.back(frame=[self.offerFrame])

    def fillOffer(self, id):
        #populates the offer frame fields with inventory item's offer details
        try:
            self.conn = MySQLConnection(**self.db)
            if self.conn.is_connected() != True:
                self._controller.log_event(self._controller.lineno(),
                                           "Database not connected contact the network admin.")
            self.c = self.conn.cursor()
            self.c.execute("SELECT On_Offer, Price_Change, Time_Limit FROM offers WHERE Item_ID=%(id)s",
                           {'id':id})
            data = self.c.fetchall()[0]
            if data:
                return {'offer':data[0], 'price':data[1], 'time':data[2]}
            else:
                return None
        except Error as e:
            self._controller.log_event(e, self._controller.lineno())
        finally:
            self.conn.commit()
            self.conn.close()

    def viewTopSelling(self):
        #sort treeview by top selling - best to worst
        self.itemTree.delete(*self.itemTree.get_children())
        try:
            self.conn = MySQLConnection(**self.db)
            if self.conn.is_connected() != True:
                self._controller.log_event(self._controller.lineno(),
                                           "Database not connected contact the network admin.")
            self.c = self.conn.cursor()
            datePast = dt.datetime.now()-dt.timedelta(days=30)
            #select all current inventory from the database ordered by income from sales
            self.c.execute("SELECT current_inventory.*, all_inventory.Batch_ID, all_inventory.Date, "
                           "inventory.Item_Name, inventory.Category, all_inventory.Original_Stock, " 
                           "all_inventory.Price_Paid, ebay_price_history.Price, offers.On_Offer, "
                           "offers.Price_Change, offers.Time_Limit, SUM(sales.Total_Income) FROM current_inventory "
                           "INNER JOIN "
                                "all_inventory ON current_inventory.Item_ID = all_inventory.Item_ID "
                           "INNER JOIN "
                                "ebay_price_history ON all_inventory.Item_ID = ebay_price_history.Item_ID "
                           "INNER JOIN "
                                "offers ON ebay_price_history.Item_ID = offers.Item_ID "
                           "INNER JOIN "
                                "inventory ON offers.Item_ID = inventory.Item_ID "
                           "LEFT JOIN sales ON (all_inventory.Batch_ID=sales.Item_Batch_ID) "
                           "WHERE all_inventory.Date_Removed IS NULL AND "
                           "all_inventory.Queue=False AND ebay_price_history.Current=True AND sales.Date >= %(date)s "
                           "GROUP BY current_inventory.Item_ID "
                           "ORDER BY SUM(sales.Total_Income)", {'date':datePast})

            self._itemHeaders = {elem[1][0]:elem[0] for elem in enumerate(self.c.description)}
            self._itemData = {elem[0]:elem[0:] for elem in self.c.fetchall()}

            #add items to treeview
            for item in self._itemData:
                self.itemTree.insert("", "end", values=[self._itemData[item][self._itemHeaders['Item_ID']],
                                                        self._itemData[item][self._itemHeaders['Item_Name']],
                                                        "£"+str(self._itemData[item][self._itemHeaders['Price_Selling']]),
                                                        int(self._itemData[item][self._itemHeaders[
                                                            'Batch_Stock']])+int(self._itemData[item][
                                                            self._itemHeaders['Total_Stock']]),
                                                        self._itemData[item][self._itemHeaders['Last_Refresh']]])

        except Error as e:
            self._controller.log_event(e, self._controller.lineno())

        finally:
            self.conn.commit()
            self.conn.close()
            self.treeChild = self.itemTree.get_children('')
            tk.messagebox.showinfo("Best to Worst", "The treeview has been re-ordered to show best to worst selling "
                                                    "products across the last MONTH. Press refresh to go back to the normal view.")

    def priceMatch(self):
        #function matches price for all items in current inventory
        confirmation = tk.messagebox.askyesno("Information", "This program will iterate through each item in your "
                                                             "current inventory and find the average value of a range "
                                                             "of prices from different eBay listings matching your past "
                                                             "search criteria. Do you want this program to check with "
                                                             "you for each item?")

        try:
            self.conn = MySQLConnection(**self.db)
            if self.conn.is_connected() != True:
                self._controller.log_event(self._controller.lineno(),
                                           "Database not connected contact the network admin.")
            self.c = self.conn.cursor()
            #gets all item and search parameter data
            self.c.execute("SELECT ebay_search_data.*, inventory.Item_Name FROM current_inventory "
                           "INNER JOIN ebay_search_data ON "
                           "current_inventory.Item_ID = ebay_search_data.Item_ID "
                           "INNER JOIN inventory ON "
                           "current_inventory.Item_ID = inventory.Item_ID")
            self._searchHeaders = {elem[1][0]:elem[0] for elem in enumerate(self.c.description)}
            self._searchData = {elem[0]:elem for elem in self.c.fetchall()}

            index = [self._searchHeaders['Min_Price'], self._searchHeaders['Max_Price'], self._searchHeaders['Item_Condition'],
                     self._searchHeaders['Listing'], self._searchHeaders['AllowReturns'],
                     self._searchHeaders['Seller_Rating'], self._searchHeaders['Business_Type'], self._searchHeaders[
                         'Search_Term']]
            for item in self._searchData:
                #iterates through each item and compares prices
                itemIdentifier = str(self._searchData[item][self._searchHeaders['Item_ID']]) + "-" + \
                                 str(self._searchData[item][self._searchHeaders['Item_Name']])
                self.parameters = {'min price': self._searchData[item][index[0]], 'max price': self._searchData[item][index[1]],
                                   'condition': self._searchData[item][index[2]], 'listing': self._searchData[item][index[3]],
                                   'returns': self._searchData[item][index[4]], 'seller rating': self._searchData[item][index[5]],
                                   'business type': self._searchData[item][index[6]], 'search': self._searchData[item][index[7]]}
                self.item = self.getEBayData(**self.parameters)
                total = 0
                for listing in self.item:
                    total += float(listing.currentprice.string)
                total /= len(self.item)
                total = round(total, 2)
                if confirmation:
                    if not tk.messagebox.askyesno(itemIdentifier, "The average for "+itemIdentifier+" is £ %.2f . Do you "
                                                  "wish to save this average? "%(total)):
                        continue
                self.c.execute("UPDATE ebay_price_history SET Current=FALSE WHERE Item_ID=%(id)s", {'id':self._searchData[item][self._searchHeaders['Item_ID']]})
                self.c.execute("INSERT INTO ebay_price_history VALUES (%(id)s, CURRENT_TIMESTAMP, %(itemID)s, "
                               "%(price)s, 'auto', TRUE )",
                               {'id':self._controller.getID(), 'itemID':self._searchData[item][self._searchHeaders['Item_ID']],
                                'price':total})


        except Error as e:
            self._controller.log_event(e, self._controller.lineno())

        finally:
            self.conn.commit()
            self.conn.close()
            #updates treeview
            self.updateItemList()

    def back(self, frame=None, update=True):
        #deletes all passed frames from display
        if frame:
            for f in frame:
                f.destroy()
        if update:
            #refreshes treeview
            self.updateItemList()

    def downloadDB(self):
        #function to download inventory details to an excel spreadsheet
        self.style0 = easyxf("align:horiz centre")
        self.style1 = easyxf('align:horiz center', num_format_str='DD-MMM-YY')
        self.style2 = easyxf('pattern: pattern solid, fore_color yellow; font: bold on;align:horiz center')

        self.book = Workbook()

        try:
            #queries all the inventory data from the database
            self.conn = MySQLConnection(**self.db)
            if self.conn.is_connected() != True:
                self._controller.log_event(self._controller.lineno(), "Database not connected contact the network "
                                                                      "admin.")
            self.c = self.conn.cursor()

            self.c.execute("SELECT all_inventory.*, inventory.Item_Name, inventory.Category "
                           "FROM all_inventory INNER JOIN inventory "
                           "WHERE all_inventory.Item_ID = inventory.Item_ID "
                           "GROUP BY all_inventory.Batch_ID ORDER BY all_inventory.Date")
            self.header = [elem[0] for elem in self.c.description]
            #iterates through each column header of the inventory database table and adds it to the spreadsheet
            self.sheet1 = self.book.add_sheet("All Inventory")
            for CI, col in enumerate(self.header):
                self.sheet1.write(0, CI, col, self.style2)
            #iterates through each inventory record to add it to the spreadsheet
            for r, row in enumerate(self.c.fetchall()):
                for c, col in enumerate(row):
                    if isinstance(col, dt.date):
                        self.sheet1.write(r + 1, c, col, self.style1)
                    else:
                        self.sheet1.write(r + 1, c, col, self.style0)

            self.c.execute("SELECT current_inventory.*, inventory.Item_Name, inventory.Category "
                           "FROM current_inventory INNER JOIN inventory "
                           "WHERE current_inventory.Item_ID = inventory.Item_ID "
                           "GROUP BY inventory.Item_ID ORDER BY current_inventory.Last_Refresh")
            self.header = [elem[0] for elem in self.c.description]
            self.sheet2 = self.book.add_sheet('Current Inventory')
            for CI, col in enumerate(self.header):
                self.sheet2.write(0, CI, col, self.style2)
            for r, row in enumerate(self.c.fetchall()):
                for c, col in enumerate(row):
                    if isinstance(col, dt.date):
                        self.sheet2.write(r + 1, c, col, self.style1)
                    else:
                        self.sheet2.write(r + 1, c, col, self.style0)

            #launch a pop-up file explorer to select a save location
            self.file = filedialog.asksaveasfilename(title="Save Inventory Data", initialdir='Log/',
                                                     defaultextension=".xls")
            if self.file:
                #if a location is selected, save the file
                self.book.save(self.file)
        except Error as e:
            self._controller.log_event(e, self._controller.lineno())

        finally:
            self.conn.close()

class RefundsPage(tk.Frame):
    def __init__(self, parent, controller, _bg):
        #creates refund page, setting values for the root window, controlling class (Main), company details,
        # background colour, font styles, user details, user ID and the company's private database
        # access details
        super().__init__(parent, bg=_bg)
        self.place(x=0, y=0, relheight=1, relwidth=1)
        self._bg = _bg
        self._top = parent
        self._controller = controller
        self.company = self._controller.get_company_metadata()
        self.user = controller.output_user()
        self.db = self._controller.getDBdetails()
        self.fonts = controller.getAppearance('font')

        self.userID = self.user['ID']

        #creates the widgets for the page that displays the sales
        self._title = tk.Label(self, text="Sales History", bg=_bg, font=self.fonts[0])
        self._title.place(relx=0.4, rely=0.01, relheight=0.05, relwidth=0.2)

        self._treeHeaders = ('Sales ID', 'Date', 'Batch ID', 'Item Name', 'Price Sold', 'Quantity',
                             'Returned')

        #treeview is like a table to display all the sales
        self.salesTree = ttk.Treeview(self, height=20, columns=self._treeHeaders, show='headings')
        for heading in self._treeHeaders:
            self.salesTree.heading(heading, text=heading, command=lambda
                col=heading: self._controller.sortItem(col, self.salesTree))
            self.salesTree.column(heading, stretch='yes', width=8)

        self.salesTree.place(relx=0.07, rely=0.1, relheight=0.75, relwidth=0.73)

        self.treeScroll = ttk.Scrollbar(self, orient="vertical", command=self.salesTree.yview)
        self.treeScroll.place(relx=0.8, rely=0.1, relheight=0.75)
        self.salesTree.configure(yscrollcommand=self.treeScroll.set)

        #binding displays the sales' details in a pop-up window when double clicked
        self.salesTree.bind("<Double-1>", lambda e: self.viewSaleDetails())

        #search element allows searching the treeview
        self.Search_ = [tk.StringVar(), tk.StringVar()]
        self._searchL = tk.Label(self, text="Search", font=self.fonts[1], bg=_bg)
        self._searchL.place(relx=0.05, rely=0.91, relwidth=0.08, relheight=0.05)
        self._searchO = ttk.OptionMenu(self, self.Search_[0], '',*self._treeHeaders)
        self._searchO.place(relx=0.13, rely=0.91, relwidth=0.09, relheight=0.05)
        self.Search_[0].set(self._treeHeaders[0])
        self._searchE = tk.Entry(self, textvariable=self.Search_[1], bg=_bg, font=self.fonts[2])
        self._searchE.place(relx=0.22, rely=0.91, relwidth=0.13, relheight=0.05)
        self._searchB = tk.Button(self, text="Go", font=self.fonts[2], bg=_bg, command=lambda: self._controller.searchItem(
            self.Search_[0].get(), self.Search_[1].get(), self.salesTree, self.treeChild), activebackground=_bg,
                                  highlightthickness=0, bd=0)
        self._searchB.place(relx=0.35, rely=0.91, relheight=0.05, relwidth=0.05)

        #binding initiates the search by clicking enter when focused on the search field
        self._searchE.bind("<Return>", lambda e: self._controller.searchItem(
            self.Search_[0].get(), self.Search_[1].get(), self.salesTree, self.treeChild))

        self.viewReturnSales_ = tk.IntVar()
        #checkbutton changes the data of the treeview to returned/partially returned sales
        self._viewReturns = tk.Checkbutton(self, text="View Returns", bg=_bg, font=self.fonts[2],
                                           variable= self.viewReturnSales_, activebackground=_bg, command=self.viewReturns)
        self._viewReturns.place(relx=0.6, rely=0.88, relheight=0.05, relwidth=0.15)

        #button initiates return for item
        self._returnsB = tk.Button(self, text="Returns", font=self.fonts[2], bg=_bg, command=self.refund)
        self._returnsB.place(relx=0.83, rely=0.1, relheight=0.05, relwidth=0.14)

        #button refreshes sales treeview
        self.IMG1 = tk.PhotoImage(file='Images/Icons/reload.gif')
        self._reloadB = tk.Button(self, image=self.IMG1, bg=self._bg, bd=0,highlightthickness=0,
                               activebackground=self._bg, command=self.viewReturns)
        self._reloadB.image = self.IMG1
        self._reloadB.place(relx=0.01, rely=0.12, relwidth=0.05, relheight=0.05)

        #button downloads all sales data
        self.IMG2 = tk.PhotoImage(file='Images/Icons/download.gif')
        self._downloadB = tk.Button(self, image=self.IMG2, bg=self._bg, bd=0,highlightthickness=0,
                               activebackground=self._bg, command=self.downloadDB)
        self._downloadB.image = self.IMG2
        self._downloadB.place(relx=0.01, rely=0.2, relwidth=0.05, relheight=0.05)

        #disables downloads is access level greater than 3
        if self.user['Access_Level'] > 3:
            self._downloadB.config(state='disabled')

        #displays YE logo
        self.IMG = tk.PhotoImage(file='Images/young enterprise 2.gif')
        self.IMG = self.IMG.subsample(2)
        self._Logo = tk.Label(self, image=self.IMG, bg=_bg)
        self._Logo.image = self.IMG
        self._Logo.place(relx=0.85, rely=0.85, relheight=0.15, relwidth=0.15)

        #update treeview of sales
        self.updateSalesList()

    def updateSalesList(self, returns=False):
        #function updates treeview of sales
        #clears sales treeview
        self.salesTree.delete(*self.salesTree.get_children())
        #gets new data and adds it to treeview
        try:
            self.conn = MySQLConnection(**self.db)
            if self.conn.is_connected() != True:
                self._controller.log_event(self._controller.lineno(),
                                           "Database not connected contact the network admin.")
            self.c = self.conn.cursor()
            if not returns:
                self.c.execute("SELECT sales.*, inventory.Item_ID, inventory.Item_Name, inventory.Category FROM sales "
                               "INNER JOIN "
                                    "all_inventory ON sales.Item_Batch_ID = all_inventory.Batch_ID "
                               "INNER JOIN "
                                    "inventory ON all_inventory.Item_ID = inventory.Item_ID "
                               "WHERE sales.Return_Quantity=0")
            else:
                self.c.execute("SELECT sales.*, inventory.Item_ID, inventory.Item_Name, inventory.Category, "
                               "returns.Return_ID, returns.Return_Date, returns.Reason, returns.Returns_Staff_ID FROM sales "
                               "INNER JOIN "
                                    "all_inventory ON sales.Item_Batch_ID = all_inventory.Batch_ID "
                               "INNER JOIN "
                                    "inventory ON all_inventory.Item_ID = inventory.Item_ID "
                               "INNER JOIN "
                                    "returns ON sales.Sale_ID = returns.Sale_ID "
                               "WHERE sales.Return_Quantity>0")


            self._salesHeadersList = [elem[0] for elem in self.c.description]
            self._salesHeaders = {elem[1]:elem[0] for elem in enumerate(self._salesHeadersList)}
            self._salesData = {elem[0]:elem for elem in self.c.fetchall()}

            for item in self._salesData:
                self.salesTree.insert("", "end", values=[self._salesData[item][self._salesHeaders['Sale_ID']],
                                                         self._salesData[item][self._salesHeaders['Date']],
                                                         self._salesData[item][self._salesHeaders['Item_Batch_ID']],
                                                         self._salesData[item][self._salesHeaders['Item_Name']],
                                                        "£" + str(self._salesData[item][self._salesHeaders['Price_Sold']]),
                                                         str(self._salesData[item][self._salesHeaders['Quantity']]),
                                                         str(self._salesData[item][self._salesHeaders[
                                                             'Return_Quantity']])])

        except Error as e:
            self._controller.log_event(e, self._controller.lineno())

        finally:
            self.conn.commit()
            self.conn.close()
            #sorts the list according to date
            self.treeChild = self.salesTree.get_children('')
            self._controller.sortItem('Date', self.salesTree, alternateSort=False)

    def refund(self):
        #function processes refunds
        if not self.salesTree.selection():
            #check if item is clicked
            tk.messagebox.showerror("Invalid Selection", "Click an item to refund.")
            return
        self.selectionID = self.salesTree.item(self.salesTree.selection(), 'values')[0]
        initial = self._salesData[self.selectionID][self._salesHeaders['Quantity']]
        refunded = self._salesData[self.selectionID][self._salesHeaders['Return_Quantity']]

        quantity = tk.simpledialog.askinteger("Quantity", "Enter refund amount: ", minvalue=1,
                                              maxvalue=initial-refunded)
        if not quantity:
            return
        reason = tk.simpledialog.askstring("Reason", "Please enter a reason for the refund.")
        if not reason or not 0 < len(reason) <= 120:
            tk.messagebox.showerror('Invalid Input',
                                    'Enter a reason for the refund no more than 120 characters.')
            return
        price = self._salesData[self.selectionID][self._salesHeaders['Price_Sold']]
        PPid = self._salesData[self.selectionID][self._salesHeaders['PayPal']]
        #check to refund with PayPal if sale has PayPal transaction ID
        if PPid:
            if tk.messagebox.askyesno("Refund with PayPal", "Do you want to refund the amount paid with PayPal?"):
                #calls the function PayPal_Refund if the user says 'yes'
                if not self.PayPal_Refund(PPid, quantity, price, self._salesData[self.selectionID][
                    self._salesHeaders[
                    'PayPal_Tax']]):
                    tk.messagebox.showerror("Refund unavailable", "The refund did not process. Please contact the "
                                                                  "system administator for further support.")
                    return

        refunded += quantity
        #alters the database to add the refund
        try:
            self.conn = MySQLConnection(**self.db)
            if self.conn.is_connected() != True:
                self._controller.log_event(self._controller.lineno(),
                                           "Database not connected contact the network admin.")
            self.c = self.conn.cursor()

            self.c.execute("UPDATE sales SET Return_Quantity=%(quantity)s, Total_Income=(Quantity-%(quantity)s)*Price_Sold "
                           "WHERE Sale_ID=%(id)s",
                           {'quantity': refunded, 'id': self.selectionID})

            self.c.execute("INSERT INTO returns VALUES (%(id)s, CURRENT_TIMESTAMP, %(sale)s, %(reason)s, %(staff)s)",
                           {'id':self._controller.getID(), 'sale':self.selectionID, 'reason':reason,
                            'staff':self.userID})

            if tk.messagebox.askyesno("Restock", "Do you wish to restock the refunded product? "):
                #if user wants to restock the product, adds it back to the inventory
                self.c.execute("SELECT Item_ID FROM all_inventory WHERE Batch_ID=%(id)s", {'id':self._salesData[self.selectionID][self._salesHeaders['Item_Batch_ID']]})
                itemID = self.c.fetchall()[0][0]
                self.c.execute("SELECT * FROM current_inventory WHERE Item_ID=%(id)s", {'id':itemID})
                if self.c.fetchall()[0]:
                    #if item is in stock currently
                    self.c.execute("INSERT INTO all_inventory VALUES (%(id)s, CURRENT_TIMESTAMP, %(itemID)s, %(stock)s, "
                                   "%(price)s, NULL, TRUE )", {'id': self._controller.getID(), 'itemID':itemID,
                                                               'stock': quantity, 'price': price})
                    self.c.execute("UPDATE current_inventory SET Total_Stock=Total_Stock+%(stock)s, "
                                   "Queue_length=Queue_length+1 WHERE Item_ID=%(id)s",
                                   {'stock': quantity, 'id': itemID})
                else:
                    #if item is no longer in stock
                    self.c.execute("INSERT INTO all_inventory VALUES (%(id)s, CURRENT_TIMESTAMP, %(itemID)s, %(stock)s, "
                                   "%(price)s, NULL, FALSE)", {'id': self._controller.getID(), 'itemID':itemID,
                                                               'stock': quantity, 'price': price})
                    self.c.execute("INSERT INTO current_inventory VALUES (%(id)s, %(stock)s, 0, %(price)s, "
                                   "CURRENT_TIMESTAMP, 0)",
                                   {'id': itemID, 'stock': quantity, 'price':price})


            tk.messagebox.showinfo("Success", "Refund successful!")
        except Error as e:
            self._controller.log_event(e, self._controller.lineno())
        finally:
            self.conn.commit()
            self.conn.close()
            #update treeview
            self.viewReturns()

    def viewReturns(self):
        #function checks the checkbutton and updates the treeview with sales or refunded sales accordingly
        if self.viewReturnSales_.get():
            self.updateSalesList(True)
        else:
            self.updateSalesList()

    def viewSaleDetails(self):
        #function creates the pop-up window to display details on the sale
        if not self.salesTree.selection():
            tk.messagebox.showerror("Error", "Please select a transaction.")
            return
        self.selectionID = self.salesTree.item(self.salesTree.selection(), 'values')[0]
        selection = self._salesData[self.selectionID]
        details = []
        self._top.clipboard_clear()
        #copies the sale ID to the system
        self._top.clipboard_append(self.selectionID)
        for a, b in zip(self._salesHeadersList, selection):
            details.append("{0}: {1}\n".format(a, b))
        tk.messagebox.showinfo("Sales ID: "+self.selectionID,"".join(details))

    def PayPal_Refund(self, id, quantity, price, tax):
        #function processes the refund through PayPal
        if tax:
            #if fee added to payment, adds fee to total amount refunded
            if quantity == 0:
                refundAmount = quantity*price*1.034
                refundAmount += 0.2
            else:
                refundAmount = quantity*price*1.034
        else:
            refundAmount = quantity*price
        refundAmount = round(refundAmount, 2)
        sale = paypal.Sale.find(id)
        # Make Refund API call
        refund = sale.refund({
            "amount": {
                "total": str(refundAmount),
                "currency": "GBP"
            }
        })

        # Check refund status
        if refund.success():
            self._controller.log_event("Refund[%s] Success" % (refund.id))
            return True
        else:
            self._controller.log_event("Unable to Refund")
            self._controller.log_event(refund.error)
            return

    def downloadDB(self):
        #function downloads sale to excel spreadsheet
        self.style0 = easyxf("align:horiz centre")
        self.style1 = easyxf('align:horiz center', num_format_str='DD-MMM-YY')
        self.style2 = easyxf('pattern: pattern solid, fore_color yellow; font: bold on;align:horiz center')

        self.book = Workbook()

        #gets refunded sales and adds it to a sheet
        self.updateSalesList(returns=True)
        self.sheet1 = self.book.add_sheet('Refunded Sales')
        for CI, col in enumerate(self._salesHeadersList):
            self.sheet1.write(0, CI, col, self.style2)
        for r, row in enumerate(list(self._salesData.values())):
            for c, col in enumerate(row):
                if isinstance(col, dt.date):
                    self.sheet1.write(r + 1, c, col, self.style1)
                else:
                    self.sheet1.write(r + 1, c, col, self.style0)

        #gets non-refunded sales and adds it to a sheet
        self.updateSalesList(returns=False)
        self.sheet2 = self.book.add_sheet('Sales')
        for CI, col in enumerate(self._salesHeadersList):
            self.sheet2.write(0, CI, col, self.style2)
        for r, row in enumerate(list(self._salesData.values())):
            for c, col in enumerate(row):
                if isinstance(col, dt.date):
                    self.sheet2.write(r + 1, c, col, self.style1)
                else:
                    self.sheet2.write(r + 1, c, col, self.style0)

        self.file = filedialog.asksaveasfilename(title="Save Sales Data", initialdir='Log/',
                                                 defaultextension=".xls")
        if self.file:
            self.book.save(self.file)

class InventoryPage(tk.Frame):
    def __init__(self, parent, controller, _bg):
        #creates inventory graphs and data page, setting values for the root window, controlling class (Main),
        # company details, background colour and font styles
        super().__init__(parent, bg=_bg)
        self.place(x=0, y=0, relheight=1, relwidth=1)
        self._bg = _bg
        self._top = parent
        self._controller = controller
        self.company = self._controller.get_company_metadata()
        self.db = self._controller.getDBdetails()
        self.fonts = controller.getAppearance('font')

        #creates the widgets involved in selecting the graph to display
        self._head1 = tk.Label(self, text="Graphs", bg=_bg, font=self.fonts[0])
        self._head1.place(relx=0.45, rely=0.05, relheight=0.05, relwidth=0.1)

        #refreshes the list of items in each optionmenu
        self.IMG1 = tk.PhotoImage(file='Images/Icons/reload.gif')
        self._reloadB = tk.Button(self, image=self.IMG1, bg=self._bg, bd=0,highlightthickness=0,
                               activebackground=self._bg, command=self.updateOptions)
        self._reloadB.image = self.IMG1
        self._reloadB.place(relx=0.6, rely=0.05, relwidth=0.05, relheight=0.05)

        #gets a list of all the options
        self.options = self.getOptions()

        #sales segment
        self._Sales = tk.StringVar()
        self._salesL = tk.Label(self, text="Sales", bg=_bg, font=self.fonts[1], justify='right', anchor='e')
        self._salesL.place(relx=0.3, rely=0.13, relheight=0.05, relwidth=0.15)
        self._salesO = ttk.OptionMenu(self, self._Sales, *self.options)
        #set sales to 'All'
        self._Sales.set(self.options[1])
        self._salesO.place(relx=0.4, rely=0.18, relheight=0.05, relwidth=0.2)
        self._salesB = tk.Button(self, text="Go", font=self.fonts[2], bg=_bg, bd=0, highlightthickness=0,
                                 command=lambda: self.createSalesGraph(self._Sales.get(), self._viewChart[0].get(),
                                                                       windowed=self._viewChart[1].get()))
        self._salesB.place(relx=0.6, rely=0.18, relheight=0.05, relwidth=0.1)

        #profit segment
        self._Profit = tk.StringVar()
        self._profitL = tk.Label(self, text="Profit", bg=_bg, font=self.fonts[1], justify='right', anchor='e')
        self._profitL.place(relx=0.3, rely=0.28, relheight=0.05, relwidth=0.15)
        self._profitO = ttk.OptionMenu(self, self._Profit, *self.options)
        #set profit to 'All'
        self._Profit.set(self.options[1])
        self._profitO.place(relx=0.4, rely=0.33, relheight=0.05, relwidth=0.2)
        self._profitB = tk.Button(self, text="Go", font=self.fonts[2], bg=_bg, bd=0, highlightthickness=0,
                                  command=lambda: self.createProfitGraph(self._Profit.get(), self._viewChart[0].get(),
                                                                        windowed=self._viewChart[1].get()))
        self._profitB.place(relx=0.6, rely=0.33, relheight=0.05, relwidth=0.1)

        #ebay segment
        self._Ebay = tk.StringVar()
        self._ebayL = tk.Label(self, text="Ebay Price History", bg=_bg, font=self.fonts[1], justify='right', anchor='e')
        self._ebayL.place(relx=0.3, rely=0.4, relheight=0.05, relwidth=0.2)
        self._ebayO = ttk.OptionMenu(self, self._Ebay, *['']+self.options[2:])
        self._ebayO.place(relx=0.4, rely=0.45, relheight=0.05, relwidth=0.2)
        self._ebayB = tk.Button(self, text="Go", font=self.fonts[2], bg=_bg, bd=0, highlightthickness=0,
                                command=lambda: self.createEbayGraph(self._Ebay.get(), self._viewChart[0].get(),
                                                                      windowed=self._viewChart[1].get()))
        self._ebayB.place(relx=0.6, rely=0.45, relheight=0.05, relwidth=0.1)
        #set eBay option to the inventory item with the lowest ID
        if len(self.options) > 2:
            self._Ebay.set(self.options[2])

        self._viewChart = [tk.StringVar(), tk.IntVar()]
        #used to adjust the x-axis intervals
        self._scaleL = tk.Label(self, text="X-Axis Scale: ", bg=_bg, font=self.fonts[2], justify='right', anchor='e')
        self._scaleL.place(relx=0.4, rely=0.6, relheight=0.05, relwidth=0.1)
        self._scaleO = ttk.OptionMenu(self, self._viewChart[0], *('','Hourly', 'Daily', 'Weekly', 'Monthly', 'Yearly'))
        self._scaleO.place(relx=0.5, rely=0.6, relheight=0.05, relwidth=0.06)
        self._viewChart[0].set("Daily")

        #used to open the graph in a new window
        self._windowB = tk.Checkbutton(self, text="New window", variable=self._viewChart[1], bg=self._bg,
                                       font=self.fonts[2])
        self._windowB.place(relx=0.43, rely=0.65, relheight=0.05, relwidth=0.1)

    def updateOptions(self):
        #updates the optionmenus menus
        self.options = self.getOptions()
        self._salesO.set_menu(*self.options)
        self._profitO.set_menu(*self.options)
        self._ebayO.set_menu(*['']+self.options[2:])

        self._Sales.set(self.options[1])
        self._Profit.set(self.options[1])
        if len(self.options) > 2:
            self._Ebay.set(self.options[2])

    def getOptions(self):
        #function gets a new list of items to display with the optionmenus
        try:
            self.conn = MySQLConnection(**self.db)
            if self.conn.is_connected() != True:
                self._controller.log_event(self._controller.lineno(),
                                           "Database not connected contact the network admin.")
            self.c = self.conn.cursor()
            self.c.execute("SELECT Item_ID, Item_Name FROM inventory")
            items = ['', 'All']+[str(elem[0])+"-"+str(elem[1]) for elem in self.c.fetchall()]
            return items

        except Error as e:
            self._controller.log_event(e, self._controller.lineno())

        finally:
            self.conn.commit()
            self.conn.close()

    def createDate(self, *date):
        #function converts dates to a matplotlib date format
        mplDt = []
        for x in date:
            mplDt.append(date2num(x))
        return mplDt

    def createSalesGraph(self, item, scale, windowed=False):
        #functions gets sale data for the selected item and displays it in a graph
        try:
            self.conn = MySQLConnection(**self.db)
            if self.conn.is_connected() != True:
                self._controller.log_event(self._controller.lineno(),
                                           "Database not connected contact the network admin.")
            self.c = self.conn.cursor()
            _itemIDName = item.split('-')
            if item == 'All':
                name = 'all'
                self.c.execute("SELECT Date, Total_Income FROM sales ORDER BY Date")
            else:
                #if selected a specific item from the optionmenu
                name = "".join(_itemIDName[1:])
                self.c.execute("SELECT sales.Date, sales.Total_Income FROM all_inventory "
                               "INNER JOIN sales ON all_inventory.Batch_ID = sales.Item_Batch_ID "
                               "WHERE all_inventory.Item_ID=%(id)s ORDER BY sales.Date", {'id': _itemIDName[0]})
            self.dataPoints = self.c.fetchall()
            if len(self.dataPoints) <= 1:
                tk.messagebox.showerror("Insufficient Data", "There is insufficient recorded data to plot a graph. "
                                                             "This usually means that 1 or no sales have been made. ")
                return
            self.showGraph(self.dataPoints, "Sales Line Graph", "Sales of "+name+" over time",
                                "Date", "Amount /£", scale, windowed=windowed)
        except Error as e:
            self._controller.log_event(e, self._controller.lineno())

        finally:
            self.conn.commit()
            self.conn.close()

    def createProfitGraph(self, item, scale, windowed=False):
        #functions gets profit data for the selected item and displays it in a graph
        try:
            self.conn = MySQLConnection(**self.db)
            if self.conn.is_connected() != True:
                self._controller.log_event(self._controller.lineno(),
                                           "Database not connected contact the network admin.")
            self.c = self.conn.cursor()
            _itemIDName = item.split('-')
            if item == 'All':
                self.c.execute("SELECT sales.Date, sales.Total_Income-(all_inventory.Price_Paid*("
                               "sales.Quantity-sales.Return_Quantity)) FROM sales INNER JOIN all_inventory "
                               "ON sales.Item_Batch_ID=all_inventory.Batch_ID "
                               "ORDER BY sales.Date")
                name = 'all'
            else:
                #if selected a specific item from the optionmenu
                name = "".join(_itemIDName[1:])
                self.c.execute("SELECT sales.Date, sales.Total_Income-(all_inventory.Price_Paid*("
                               "sales.Quantity-sales.Return_Quantity)) FROM all_inventory "
                               "INNER JOIN sales ON all_inventory.Batch_ID = sales.Item_Batch_ID "
                               "WHERE all_inventory.Item_ID=%(id)s ORDER BY sales.Date", {'id': _itemIDName[0]})
            self.dataPoints = self.c.fetchall()
            if len(self.dataPoints) <= 1:
                tk.messagebox.showerror("Insufficient Data", "There is insufficient recorded data to plot a graph. "
                                                             "This usually means that 1 or no sales have been made. ")
                return

            self.showGraph(self.dataPoints, "Profit Line Graph", "Profit of "+name+" over time",
                                "Date", "Amount /£", scale, windowed=windowed)
        except Error as e:
            self._controller.log_event(e, self._controller.lineno())

        finally:
            self.conn.commit()
            self.conn.close()

    def createEbayGraph(self, item, scale, windowed=False):
        #functions gets ebay graph data for the selected item and displays it in a graph
        try:
            self.conn = MySQLConnection(**self.db)
            if self.conn.is_connected() != True:
                self._controller.log_event(self._controller.lineno(),
                                           "Database not connected contact the network admin.")
            self.c = self.conn.cursor()
            _itemIDName = item.split('-')
            self.c.execute("SELECT Date, Price FROM ebay_price_history WHERE ebay_price_history.Item_ID=%(id)s "
                           "ORDER BY Date", {'id': _itemIDName[0]})
            self.dataPoints = self.c.fetchall()
            if len(self.dataPoints) <= 1:
                tk.messagebox.showerror("Insufficient Data", "There is insufficient recorded data to plot a graph. "
                                                             "This usually means that only 1 eBay price has been "
                                                             "recorded. Click the Match Price button in the Buying "
                                                             "page of the Inventory menu to automatically add another piece of price data. ")
                return


            self.showGraph(self.dataPoints, "eBay Price Track Line Graph", "eBay prices of "+
                           "".join(_itemIDName[1:])+" over time", "Date", "Amount /£", scale, windowed=windowed)
        except Error as e:
            self._controller.log_event(e, self._controller.lineno())

        finally:
            self.conn.commit()
            self.conn.close()

    def showGraph(self, dataPoints, label, title, xAxis, yAxis, scale, windowed=False):
        #function displays the graph
        if not windowed:
            #if new window checkbutton not pressed, it creates a new frame in the current window
            self.ChartFrame = tk.LabelFrame(self, text=label, bg=self._bg)
            self.ChartFrame.place(relx=0.03, rely=0.03, relheight=0.94, relwidth=0.94)
            self.windowedChart = False
        else:
            #if new window checkbutton pressed, it creates a new window to display the graph
            self.ChartFrame = tk.Toplevel(bg=self._bg)
            self.ChartFrame.geometry('700x700')
            self.windowedChart = True

        #turns all x-axis values into an array
        xAxisValues = NPArray(self.createDate(*[elem[0] for elem in dataPoints]))
        #sets the scale of the graph. If it is too small, output an error message
        timeDelta = num2date(xAxisValues.max()) - num2date(xAxisValues.min())
        if scale == 'Hourly':
            rule = rrulewrapper(HOURLY, interval=1)
            if timeDelta > dt.timedelta(days=20):
                self.closeChart(full=False)
                tk.messagebox.showerror("Interval too short", "This interval is too short to display all the data "
                                                              "since this data set currently spans a date range of "
                                                              "greater than 20 days. Please choose a greater interval.")
                return
        elif scale == 'Daily':
            rule = rrulewrapper(DAILY, interval=1)
            if timeDelta > dt.timedelta(days=365):
                self.closeChart(full=False)
                tk.messagebox.showerror("Interval too short", "This interval is too short to display all the data "
                                                              "since this data set currently spans a date range of "
                                                              "greater than 365 days. Please choose a greater "
                                                              "interval.")
                return
        elif scale == 'Weekly':
            rule = rrulewrapper(WEEKLY, interval=1)
            if timeDelta > dt.timedelta(days=3360):
                self.closeChart(full=False)
                tk.messagebox.showerror("Interval too short", "This interval is too short to display all the data "
                                                              "since this data set currently spans a date range of "
                                                              "greater than 3360 days. Please choose a greater "
                                                              "interval.")
                return
        elif scale == 'Monthly':
            rule = rrulewrapper(MONTHLY, interval=1)
            if timeDelta > dt.timedelta(days=14400):
                self.closeChart(full=False)
                tk.messagebox.showerror("Interval too short", "This interval is too short to display all the data "
                                                              "since this data set currently spans a date range of "
                                                              "greater than 14400 days. Please choose a greater "
                                                              "interval.")
                return
        elif scale == 'Yearly':
            rule = rrulewrapper(YEARLY, interval=1)
            if timeDelta > dt.timedelta(days=175200):
                self.closeChart(full=False)
                tk.messagebox.showerror("Interval too short", "This interval is too short to display all the data "
                                                              "since this data set currently spans a date range of "
                                                              "greater than 175200 days. Please choose a greater "
                                                              "interval.")
                return

        #creates the plot
        fig = plt.Figure(figsize=(5, 4), dpi=200)
        fig1, ax = plt.subplots()
        plt.subplots_adjust(bottom=0.13)
        line1 = NPArray([elem[1] for elem in dataPoints])

        #sets graph attributes
        ax.plot_date(xAxisValues, line1, fmt='b-')
        ax.axis('tight')
        ax.grid(color='g', linestyle=':')
        ax.xaxis_date()
        loc = RRuleLocator(rule)
        formatter = DateFormatter("%d '%b (%H:%M)")

        ax.xaxis.set_major_locator(loc)
        ax.xaxis.set_major_formatter(formatter)
        labelsx = ax.get_xticklabels()
        plt.setp(labelsx, rotation=30, fontsize=7)

        ax.yaxis.set_major_formatter(FormatStrFormatter('%.2f'))

        #sets the graphs axis and title labels
        plt.xlabel(xAxis)
        plt.ylabel(yAxis)
        plt.suptitle(title)

        #puts the graph on the page
        plt.grid(True)

        #adds the toolbar to the bottom of the graph
        self._chartCanvas = FigureCanvasTkAgg(fig1, self.ChartFrame)
        self._chartCanvas.draw()
        self._chartCanvas.get_tk_widget().place(relx=0, rely=0, relheight=0.94, relwidth=0.9)

        self._chartBar = NavigationToolbar2Tk(self._chartCanvas, self.ChartFrame)
        self._chartBar.update()
        self._chartCanvas._tkcanvas.place(relx=0, rely=0, relheight=0.94, relwidth=0.9)

        #creates a button to close the graph if clicked
        self._closeB = tk.Button(self.ChartFrame, text="X", bg=self._bg, font=self.fonts[5], command=self.closeChart,
                                 bd=0, highlightthickness=0)
        self._closeB.place(relx=0.94, rely=0.01, relheight=0.05, relwidth=0.05)

    def closeChart(self, full=True):
        #function closes the graph. The full parameter exists because if the user chooses a too small interval,
        # it needs a way to close the frame, but the graph has not been created yet
        if full:
            self._chartCanvas.get_tk_widget().destroy()
        self.ChartFrame.destroy()

if __name__ == "__main__":
    #imports some of the main functions
    import datetime as dt
    import inspect
    import logging
    from os import listdir

    from smtplib import SMTP
    from string import ascii_uppercase, ascii_lowercase, digits
    import webbrowser as webb
    from configparser import ConfigParser
    from email.mime.multipart import MIMEMultipart
    from email.mime.text import MIMEText
    from random import choice

    from ebaysdk.finding import Connection as eBayConn
    import paypalrestsdk as paypal
    from bs4 import BeautifulSoup
    from mysql.connector import MySQLConnection, Error

    #creates the root tkinter window and raises it to the top
    root = tk.Tk()
    root.lift()

    #imports rest of modules
    from xlwt import easyxf, Workbook
    from matplotlib import use as MPLUse
    MPLUse('TkAgg')
    from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg, NavigationToolbar2Tk
    import matplotlib.pyplot as plt
    from matplotlib.dates import HOURLY, DAILY, WEEKLY, MONTHLY, YEARLY, DateFormatter, rrulewrapper, RRuleLocator, \
                                                                      date2num, num2date
    from matplotlib.ticker import FormatStrFormatter
    from numpy import array as NPArray
    from plotly.offline import plot
    from plotly.figure_factory import create_gantt

    #instantiates the program
    app = Main(root)
    root.mainloop()