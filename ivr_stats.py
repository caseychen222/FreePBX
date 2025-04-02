import PySimpleGUI as sg
import pymysql as psql
from datetime import datetime

class IVRStatistics:
    def __init__(self):
        self.callerleft = 0
        self.ivrterm = 0
        self.schools = ["All Locations"]
        self.colwidths = [10, 17, 11, 11, 8, 15]
        self.headings = ["Location:", "Timestamp:", "Caller:", "IVR's DID:", "Call Length:", "End Result:"]
        self.stats = []
        self.filtered_stats = []
        self.sort_column = None
        self.sort_reverse = False
        self.sqljoin = """SELECT
						asterisk.ivr_details.description AS ivrname,
						ivrid,
						tstamp,
						caller,
						callednum,
						calllength,
						CASE
							WHEN strcmp(sq2.selected, 's') = 0
								THEN 'Released by caller'
							WHEN strcmp(sq2.selected, 't') = 0
								THEN 'Released by IVR'
							WHEN char_length(sq2.selected) = 4
								THEN concat(ie.selection, ' -> ', sq2.selected)
							ELSE
								concat(ie.selection, ' -> ', substr(ie.dest, 17, 4))
						END AS endresult
						FROM (
							SELECT DISTINCT
								calldate AS tstamp,
								substr(asterisk.incoming.destination, 5, 1) AS ivrid,
								asterisk.incoming.extension AS ivrext,
								src AS caller,
								did AS callednum,
								dst AS selected,
								billsec AS calllength

							FROM asteriskcdrdb.cdr
							
							JOIN
								asterisk.incoming ON asterisk.incoming.extension = did
						)sq2

					JOIN
						asterisk.ivr_entries ie ON (substr(ie.dest, 17, 4) = sq2.selected OR
						char_length(sq2.selected) = 1) AND
						ie.ivr_id = sq2.ivrid
					JOIN
						asterisk.ivr_details ON asterisk.ivr_details.id = sq2.ivrid
						
					WHERE
						sq2.callednum IN (ivrext) 

					GROUP BY tstamp
					ORDER BY tstamp DESC"""
        

        self.connection = psql.connect(
                host = "HOSTNAME",
                user = "USERNAME",
                password = "PASSWORD",
                db = "asterisk",
                )
        self.cur = self.connection.cursor()

        sg.theme("Reddit")

        self.layout = self.create_layout()

        self.window = sg.Window("FreePBX IVR Statistics", self.layout, use_default_focus=False, size = (700, 650), finalize=True)
  
    def create_layout(self):
        trow = [
            [[sg.Text("Select a location:")],
            [sg.Combo(tuple(self.schools), expand_x = True, enable_events = True, readonly = True, key = "-SCHOOLLIST-")],
            [
                [
                    sg.Input("Start:", key = "-TXTSTART-", size = (15, 0), disabled = True, border_width = None, disabled_readonly_background_color = None, enable_events = True),
                    sg.Input("End:", key = "-TXTEND-", size = (15, 0), disabled = True, border_width = None, disabled_readonly_background_color = None, enable_events = True),
                    sg.Text("Total:", key = "-TOTALCALLS-", size = (15, 0), justification = "right"),
                    sg.Text("Abandoned:", key = "-CALLERCALLS-", size = (15, 0), justification = "right"),
                    sg.Text("Kicked off:", key = "-IVRCALLS-", size = (15, 0), justification = "right")],
            ]
            ]]
        tablerow = [
            [[sg.Table(self.stats, key = "-STATS-", headings = self.headings, auto_size_columns = False, display_row_numbers = False, col_widths = self.colwidths, size = (None, 30))]]
        ]
        btnrow = [
            [[sg.Button("Filter by date"),
              sg.Button("Clear filter"),
              sg.CalendarButton(button_text = "Start date", key = "-CALSTART-", target = "-TXTSTART-", format = f"Start: %Y-%m-%d", enable_events = True),
              sg.CalendarButton(button_text = "End date", key = "-CALEND-", target = "-TXTEND-", format = f"End: %Y-%m-%d", enable_events = True)]]
        ]

        layout = [
            [trow],
            [tablerow],
            [btnrow]
            ]

        return layout
    
    def sec_to_hms(self, seconds):
        return f"{seconds // 3600:02d}:{(seconds % 3600) // 60:02d}:{seconds % 60:02d}"
    
    def populate_menu(self):
        try:
            self.cur.execute("SELECT description FROM asterisk.ivr_details WHERE description NOT LIKE '%IVR%'")
            output = self.cur.fetchall()

            for item in output:
                school = str(item[0])
                self.schools.append(school)

                self.window['-SCHOOLLIST-'].update(values = self.schools)

            self.load_statistics()

        except psql.Error as e:
            self.status_window(f"In load_statistics()\n\nError: {e}")

    def update_stats(self, data):
        total_calls = len(data)
        caller_disconnected = len([row for row in data if row[5] == 'Released by caller'])
        ivr_disconnected = len([row for row in data if row[5] == 'Released by IVR'])
        
        self.window["-STATS-"].update(values=data)
        self.window["-TOTALCALLS-"].update(value=f"Total: {total_calls}")
        self.window["-CALLERCALLS-"].update(value=f"Abandoned: {caller_disconnected}")
        self.window["-IVRCALLS-"].update(value=f"Kicked off: {ivr_disconnected}", text_color = ("red" if ivr_disconnected > 0 else "black"))

    def format_number(self, num):
        return f"({num[-10:-7]}) {num[-7:-4]}-{num[-4:]}" if len(num) >= 10 else num

    def load_statistics(self):
        try:
            self.cur.execute(self.sqljoin)
            output = self.cur.fetchall()

            for info in output:
                schname, _, tstamp, caller, did, seconds, result = info
                tstamp_fmt = tstamp.strftime("%Y-%m-%d %I:%M:%S %p")

                self.stats.append([schname, tstamp_fmt, self.format_number(caller), self.format_number(did), self.sec_to_hms(seconds), result])

            self.update_stats(self.stats)

        except psql.Error as e:
            self.status_window(f"In load_statistics()\n\nError: {e}")

    def filter_ivr(self, selected, sdate, edate):      
        if selected == sdate == edate == None:
            self.filtered_stats = self.stats

            self.update_stats(self.filtered_stats)
            self.window["-STATS-"].update(values = self.filtered_stats)

            return

        if ": " in sdate or ": " in edate:
            new_sdate = datetime.strptime(sdate[-10:], "%Y-%m-%d")
            new_edate = datetime.strptime(edate[-10:], "%Y-%m-%d")

            if new_edate < new_sdate:
                self.status_window("Start date must come after end date.")

                return

            if selected == "All Locations" or selected == None:
                self.update_stats([row for row in self.stats if (datetime.strptime(row[1][:10], "%Y-%m-%d") >= new_sdate and
                                                                 datetime.strptime(row[1][:10], "%Y-%m-%d") <= new_edate)])

                return

            elif selected and selected != "All Locations":
                self.filtered_stats = [row for row in self.stats if ((datetime.strptime(row[1][:10], "%Y-%m-%d") >= new_sdate and
                                                                      datetime.strptime(row[1][:10], "%Y-%m-%d") <= new_edate) and
                                                                      row[0] == selected)]

            self.update_stats(self.filtered_stats)
            self.window["-STATS-"].update(values = self.filtered_stats)

        else:
            if selected == "All Locations" or selected == None:
                self.update_stats(self.stats)

                return
            
            elif selected and selected != "All Locations":
                self.filtered_stats = [row for row in self.stats if row[0] == selected]

        self.update_stats(self.filtered_stats)
        self.window["-STATS-"].update(values = self.filtered_stats)

    def status_window(self, message):
        layout = [
            [sg.Text(message, justification = "center")],
            [sg.HSeparator()],
            [sg.Button("OK", bind_return_key=True)]
            ]
    
        statuswin = sg.Window("Warning", layout, element_justification = "center", finalize=True, modal=True, grab_anywhere = False)

        event, values = statuswin.read()

        statuswin.close()

    def run(self):
        self.populate_menu()
        startflag = endflag = selection = False

        while True:
            event, values = self.window.read()

            if event == sg.WIN_CLOSED:
                break

            elif event == "Clear filter":
                self.filter_ivr(None, None, None)
                self.window["-TXTSTART-"].update("Start:")
                self.window["-TXTEND-"].update("End:")
                self.window["-SCHOOLLIST-"].update(set_to_index=[])
                startflag = endflag = False

            elif event == "-SCHOOLLIST-":
                selection = True

                self.filter_ivr(values["-SCHOOLLIST-"], values["-TXTSTART-"], values["-TXTEND-"])

            elif event == "-TXTSTART-":
                startflag = True

            elif event == "-TXTEND-":
                endflag = True

            elif event == "Filter by date":
                if startflag and endflag:
                    self.filter_ivr(None if not selection else values["-SCHOOLLIST-"], values["-TXTSTART-"], values["-TXTEND-"])
                else:
                    self.status_window("Please choose a start and end date to\nfilter statistics on.")

        self.connection.close()
        self.window.close()

if __name__ == "__main__":
    ivrstats = IVRStatistics()
    ivrstats.run()
