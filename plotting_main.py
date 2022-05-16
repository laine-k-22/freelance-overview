import sys

import pandas as pd
from pandas_ods_reader import read_ods
import matplotlib.pyplot as plt
import matplotlib.dates as mdates
import seaborn as sns
import tkinter as tk
from tkinter import filedialog as fd
from tkinter import messagebox
from ctypes import windll
from PIL import ImageTk, Image
import webbrowser

file_path = ""

# Create tkinter window.
root = tk.Tk()
root.title("Freelance Overview")

window_width = 460
window_height = 480

# Get the screen dimension
screen_width = root.winfo_screenwidth()
screen_height = root.winfo_screenheight()

# Find the center point
center_x = int(screen_width / 2 - window_width / 2)
center_y = int(screen_height / 2 - window_height / 2)

# Set the position of the window to the center of the screen
root.geometry(f'{window_width}x{window_height}+{center_x}+{center_y}')
root.resizable(False, False)
root.iconbitmap("./assets/Icon_01.ico")

# Message
tk_message = "This program analyzes Open Office Spreadsheets.\n" \
             "" \
             "\nThe Spreadsheet must be organised in a specific way, \n" \
             "for more guidance on this, please click on graph below."
message = tk.Label(root, text=tk_message, font=("Arial", 10))
message.pack(pady=(40, 0))


# Define a website call function
def callback(url):
    webbrowser.open_new_tab(url)


logo = ImageTk.PhotoImage(Image.open("./assets/Logo_01.png"))
git_logo = ImageTk.PhotoImage(Image.open("./assets/Github_logo_01.png"))
graph_img = ImageTk.PhotoImage(Image.open("./assets/Graph_01.png"))

label1 = tk.Label(root, image=logo, cursor="hand2")
label1.place(relx=.95, rely=.95, anchor="se")
label1.bind("<Button-1>", lambda e: callback("http://www.lainekocane.com"))

label2 = tk.Label(root, image=git_logo, text="Laine K, 2022", compound="bottom", cursor="hand2")
label2.bind("<Button-1>", lambda e: callback("https://github.com/laine-k-22"))
label2.place(relx=.1, rely=.95, anchor="sw")

label3 = tk.Label(root, image=graph_img, cursor="hand2")
label3.bind("<Button-1>", lambda e: callback("https://github.com/laine-k-22/freelance-overview/tree/main/spreadsheet"))
label3.place(relx=.5, rely=.5, anchor="center")


# Get Spreadsheet file path and close GUI.
def select_file():
    filetypes = (
        [("OpenOffice files", ".ods")]
    )

    filename = fd.askopenfilename(
        title="Select Open Office Spreadsheet",
        initialdir="/",
        filetypes=filetypes)
    global file_path
    file_path = filename
    root.destroy()


tk.Button(root, text="Select Open Office Spreadsheet to Start!", command=select_file,
          font=("Arial", 10)).pack(pady=15)

windll.shcore.SetProcessDpiAwareness(1)
root.mainloop()

# End program if GUI is closed or file not selected.
if not file_path:
    sys.exit()

# Continue plotting the graphs.
df = read_ods(file_path, sheet=1)

# End program if column count appears wrong.
if len(df.columns) != 5:
    messagebox.showerror("Error", "Spreadsheet incorrect, please try again.")
    sys.exit()

df.columns = ["Date", "Studio", "Type", "Rate", "WFH"]
df.dropna(axis=0, how="any", thresh=4, inplace=True)

# End program if df is empty.
if df.empty:
    messagebox.showerror("Error", "Spreadsheet incorrect, please try again.")
    sys.exit()

df["Studio"] = df["Studio"].str.strip()

# Return the 2 years recorded in the current Spreadsheet.
df["Date"] = pd.to_datetime(df["Date"])

# Get years recorded from Spreadsheet.
year_sums = df["Date"].dt.year
year_sums = year_sums.groupby(year_sums).size()
year_sums = year_sums.index.tolist()

if int(len(year_sums)) == 1:
    c_year = str(year_sums[0])
else:
    c_year = str(year_sums[0]) + "-" + str(year_sums[1])

business_days = 256

# Create the figure.
gridsize = (4, 3)
fig = plt.figure(figsize=(16, 9), num="Freelance Overview Charts")


def days_worked():
    # Find how many days worked overall in the business year.
    worked_total = df["Date"].count()

    # Longest consecutive bookings.
    aggfunc = {
        "Studio": [("Studio", "first"), ("Days", "count")],
        "Type": [("Type", "first")]
    }

    grouper = df["Studio"].ne(df["Studio"].shift()).cumsum()

    lon = df.assign(key=grouper).groupby("key").agg(aggfunc)
    lon.columns = lon.columns.droplevel(0)

    # Find longest booking based on days.
    lc = lon.nlargest(n=2, columns="Days", keep="all")

    longest = lc.nlargest(n=1, columns="Days", keep="all")
    longest_days = int(longest["Days"])
    longest_studio = longest["Studio"].to_string(index=False, header=False)
    sec_longest = lc[1:2]
    second_days = int(sec_longest["Days"])
    second_studio = sec_longest["Studio"].to_string(index=False, header=False)

    # Plot a "Horizontal Fill bar".
    ax1 = plt.subplot2grid(gridsize, (0, 0), colspan=1, rowspan=1)
    ax1.set_title("Days worked in the " + c_year + " Tax Year.", fontsize=10, )
    ax1.barh([1, 1], [business_days], height=1, color="powderblue", alpha=0.15, label="Business Days.")
    # Plotting days worked bar over the previous bar.
    ax1.barh([1, 1], [worked_total], height=1, color="springgreen", label="Days Worked.")
    ax1.barh([1, 1], [longest_days], height=1, color="darkturquoise", alpha=1,
             label=f"Longest single consecutive booking.\nStudio: {longest_studio}")
    ax1.barh([1, 1], [second_days], height=1, color="royalblue",
             label=f"Second longest. Studio: {second_studio}")

    # Set Y axis row count.
    ax1.set_ylim(0, 3)
    # Hide X and Y axis.
    ax1.get_xaxis().set_visible(False)
    ax1.get_yaxis().set_visible(False)

    ax1.legend(bbox_to_anchor=(1, 1), bbox_transform=fig.transFigure)
    ax1.legend(bbox_to_anchor=(0, .5, 1.1, .5), loc="upper left", ncol=2, mode="expand",
               borderaxespad=0., frameon=False, fontsize="small")
    # Hide frame.
    for pos in ["right", "top", "bottom", "left"]:
        plt.gca().spines[pos].set_visible(False)

    # Adding labels.
    ax1.bar_label(ax1.containers[0], label_type="edge", color="black", fontsize=8, padding=3)
    ax1.bar_label(ax1.containers[1], label_type="edge", color="black", fontsize=8, padding=3)
    ax1.bar_label(ax1.containers[2], label_type="edge", color="black", fontsize=8, padding=3)
    ax1.bar_label(ax1.containers[3], label_type="edge", color="black", fontsize=8, padding=3)


# Find total profit and deduce approx. tax.
gross_profit = df["Rate"].sum()
net_profit = gross_profit - (gross_profit * 0.2)

income_total = f"Gross Profit : £ {gross_profit:,.2f}\n" \
               f"Net Profit estimate (after deducing 20% Tax): £ {net_profit:,.2f}"


def most_least_month():
    # Group dates by months and calculate monthly earnings.

    # Get days worked in each month and add them in new column.
    df["Count_Days"] = df.groupby([pd.Grouper(key="Date", freq="M")])["Date"].transform("size").astype(int)

    df.set_index("Date", inplace=True)

    # Combine days into months and sum up monthly earnings and days worked in each month.
    md = df["Rate"].resample("M").sum()
    dc = df["Count_Days"].resample("M").mean()

    join_df = pd.concat([md, dc], axis=1, join="outer")

    # Create Bar graph.

    # Initial bar graph was starting with the wrong month.
    # Creating a new column and changing its date time format fixed this.
    join_df["Plot Dates"] = join_df.index
    join_df["Plot Dates"] = pd.to_datetime(join_df["Plot Dates"].dt.strftime("%Y-%m"))

    # Set graph size.
    ax2 = plt.subplot2grid(gridsize, (2, 0), colspan=1, rowspan=2)

    ax2.bar(join_df["Plot Dates"],
            join_df["Rate"],
            width=25,
            color="orange",
            align="center",
            alpha=0.7)

    # Styling the graph.
    ax2.set(xlabel=None,
            ylabel=None)

    ax2.spines["top"].set_visible(False)
    ax2.spines["right"].set_visible(False)
    ax2.spines["left"].set_visible(False)
    ax2.spines["bottom"].set_color("#DDDDDD")

    ax2.tick_params(bottom=True, left=False, color="peru", labelsize=8)

    ax2.set_axisbelow(True)
    ax2.yaxis.grid(True, color="wheat")
    ax2.xaxis.grid(False)

    # Define the date format
    date_form = mdates.DateFormatter("%m/%Y")
    ax2.xaxis.set_major_formatter(date_form)

    # Adding labels.
    ax2.bar_label(ax2.containers[0], label_type="edge", color="peru", fontsize=8, padding=2, alpha=0.6)

    ax2.set_title("Monthly profits in given currency.", fontsize=10)
    ax2.set_facecolor("whitesmoke")

    # Reset the index so that it does not cause problems in other functions.
    df.reset_index(inplace=True)


def most_type():
    # Find which type of work is more popular.
    work_types = df["Type"].value_counts()

    # Plot a pie chart.
    colors = sns.color_palette("YlOrRd")

    ax3 = plt.subplot2grid(gridsize, (0, 2), colspan=1, rowspan=2)

    ax3.pie(work_types, labels=work_types.index, colors=colors,
            autopct='%1.0f%%', pctdistance=0.73, startangle=45,
            textprops={"fontsize": 8})

    my_circle = plt.Circle((0, 0), 0.5, color="white")
    p = plt.gcf()
    p.gca().add_artist(my_circle)
    ax3.set_title("Work Type Overview.", fontsize=10)


def most_least_bookings():
    # Find studios with most bookings in total days summed up for whole year
    # and studios with least.

    # Group bookings from same studio and count total booked days.
    st = df.pivot_table(columns=["Studio"], aggfunc="size")

    st = pd.DataFrame(st)
    st.reset_index(inplace=True)
    st.columns = ["Studio", "Days"]
    st = st.groupby(st["Studio"]).sum().reset_index()

    # Group Studios with up to 3 day bookings together.
    def grouping(dtf, ind, col):
        if dtf[col].loc[ind] <= 4:
            return '"Other"'

    summed_least = st.groupby(lambda x: grouping(st, x, "Days")).sum().reset_index()
    summed_least.columns = ["Studio", "Days"]

    # Make separate dataframe with only bookings of more than 3 days.
    summed_most = st[st["Days"] >= 5]

    # Join the 2 new dataframes into one.
    clean_st = pd.concat([summed_most, summed_least], axis=0, join="inner")
    clean_st = clean_st.sort_values(by="Days", ascending=False).reset_index(drop=True)

    # Plot the pie chart.
    clean_st.sort_values("Days", inplace=True)
    colors = sns.color_palette("summer_r", len(clean_st["Days"]), desat=1)
    ax4 = plt.subplot2grid(gridsize, (0, 1), colspan=1, rowspan=2)
    ax4.pie(clean_st["Days"], labels=clean_st["Studio"], colors=colors,
            autopct='%1.0f%%', pctdistance=0.73, startangle=100, rotatelabels=0,
            textprops={"fontsize": 8})

    my_circle = plt.Circle((0, 0), 0.5, color="white")
    s = plt.gcf()
    s.gca().add_artist(my_circle)
    plt.title("Amount of bookings in days for each Studio.", fontsize=10)

    dfs = st[st["Days"] <= 4]

    # Add text of 'Other' studios if there's any.
    if not dfs.empty:
        other_b = f'"Other" Category\nof shortest bookings:\n\n{dfs.to_string(index=False, header=True)}'
        ax4.text(s=other_b, x=1.7, y=0.0, family="monospace", fontsize=8)


def lon_cons():
    # Counting consecutive same string values.

    aggfunc = {
        "Studio": [("Studio", "first"), ("Days", "count")],
        "Type": [("Type", "first")]
    }

    grouper = df["Studio"].ne(df["Studio"].shift()).cumsum()

    lon = df.assign(key=grouper).groupby("key").agg(aggfunc)
    lon.columns = lon.columns.droplevel(0)

    # Find longest booking based on days.
    lc = lon.nlargest(n=6, columns="Days")

    # Plot the data.
    ax9 = plt.subplot2grid(gridsize, (1, 0), colspan=1, rowspan=1)
    ax9.bar(lc["Studio"], lc["Days"], color="red", alpha=0.5)

    # Adding labels.
    ax9.bar_label(ax9.containers[0], label_type='edge', color="black", fontsize=8, padding=2)

    ax9.set_title("5 Studios with longest consecutive bookings in Days.", fontsize=10, )


def pop_date():
    # Date that is most often booked on and which studio does the booking.
    # Also day of the week most often booked on.

    # Separate the Days from DateTime.
    df["Day"] = df["Date"].dt.day

    # Group days and count their frequency.
    d_count = df.pivot_table(index=["Day"], aggfunc="size")
    d_count = pd.DataFrame(d_count)
    d_count.reset_index(inplace=True)

    d_count.columns = ["Day", "Count"]

    # Get most frequently booked weekday.
    df["Weekday"] = df["Date"].dt.day_name()

    wd = df.pivot_table(index=["Weekday"], aggfunc="size")
    wd = pd.DataFrame(wd)
    wd.reset_index(inplace=True)
    wd.columns = ["Weekday", "Times"]

    # Create Bar graph for bookings on each date of the month.

    # Set graph size.
    ax5 = plt.subplot2grid(gridsize, (2, 1), colspan=1, rowspan=2)

    # Add x-axis and y-axis
    ax5.bar(d_count["Day"],
            d_count["Count"],
            color="palegreen",
            alpha=0.7)

    # Style the graph.
    ax5.spines["top"].set_visible(False)
    ax5.spines["right"].set_visible(False)
    ax5.spines["left"].set_visible(False)
    ax5.spines["bottom"].set_color("#DDDDDD")

    ax5.tick_params(bottom=True, left=False, color="greenyellow", labelsize=8)

    ax5.set_axisbelow(True)
    ax5.yaxis.grid(True, color="greenyellow")
    ax5.xaxis.grid(False)

    # Set title and labels for axes
    ax5.set(xlabel="Date",
            ylabel="Frequency")

    ax5.set_title("Frequency of bookings\nfor each day of the month.", fontsize=10)
    ax5.set_facecolor("whitesmoke")

    # Create Bar graph for bookings on each day of the week.

    # Organise weekdays in the dataframe.
    cats = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"]
    wd = wd.set_index("Weekday").reindex(cats).reset_index()

    wd["Weekday"] = wd["Weekday"].replace(["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday",
                                           "Sunday"], ["Mon", "Tues", "Wedn", "Thurs", "Fri", "Sat", "Sun"])

    # Plot graph.
    ax6 = plt.subplot2grid(gridsize, (2, 2), colspan=1, rowspan=2)

    # Add x-axis and y-axis
    ax6.bar(wd["Weekday"],
            wd["Times"],
            color="deepskyblue",
            alpha=0.5)

    # Style the graph.
    ax6.spines["top"].set_visible(False)
    ax6.spines["right"].set_visible(False)
    ax6.spines["left"].set_visible(False)
    ax6.spines["bottom"].set_color("#DDDDDD")

    ax6.tick_params(bottom=True, left=False, color="paleturquoise", labelsize=8)

    ax6.set_axisbelow(True)
    ax6.yaxis.grid(True, color="paleturquoise")
    ax6.xaxis.grid(False)

    # Set title and labels for axes
    ax6.set(xlabel="Day of the Week",
            ylabel="Frequency")

    # Adding labels.
    ax6.bar_label(ax6.containers[0], label_type="edge", color="deepskyblue", fontsize=8, padding=2, alpha=0.45)
    ax6.set_title("Frequency of bookings\nfor each day of the Week.", fontsize=10)
    ax6.set_facecolor("whitesmoke")


def wfh():
    # Returns sum of days worked from home and sum of worked in-house.
    work_f_home = df["WFH"].sum()

    inh = 0

    for i in df["WFH"]:
        if i == 0:
            inh = inh + 1

    # Plot a "Horizontal Fill bar".
    ax7 = plt.subplot2grid(gridsize, (1, 0), colspan=1, rowspan=1)
    ax7.set_title("Days worked from Home in the " + c_year + " Tax Year.", fontsize=10, )
    ax7.barh([1, 1], [work_f_home], height=1, color="greenyellow", alpha=0.40, label="Worked from Home.")
    # Plotting days worked bar over the previous bar.
    ax7.barh([1, 1], [inh], height=1, color="darkorange", label="Worked On-Site.")

    # Set Y axis row count.
    ax7.set_ylim(0, 3)
    # Hide X and Y axis.
    ax7.get_xaxis().set_visible(False)
    ax7.get_yaxis().set_visible(False)
    ax7.legend(frameon=False, loc="upper right", fontsize="small")

    # Hide frame.
    for pos in ["right", "top", "bottom", "left"]:
        plt.gca().spines[pos].set_visible(False)

    # Adding labels.
    ax7.bar_label(ax7.containers[0], label_type="edge", color="black", fontsize=8, padding=3)
    ax7.bar_label(ax7.containers[1], label_type="edge", color="black", fontsize=8, padding=3)


days_worked()
most_least_month()
most_type()
most_least_bookings()
pop_date()
wfh()

# Adding figure title and text.
fig.suptitle(c_year + " Freelance Year Overview", fontsize=14)
fig.subplots_adjust(top=0.80)
fig.text(s=income_total, x=0.12, y=0.88, fontsize=12, ha="left", va="baseline")
fig.subplots_adjust(wspace=0.236, hspace=0.55)
# Show figure in full screen.
manager = plt.get_current_fig_manager()
manager.resize(*manager.window.maxsize())
manager.window.wm_iconbitmap("./assets/Icon_01.ico")

plt.show()

sys.exit()
