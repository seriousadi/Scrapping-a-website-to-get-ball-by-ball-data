# imports for program
from selenium import webdriver
from selenium.webdriver.common.by import By
from bs4 import BeautifulSoup
import pandas as pd
import time

# website to parse
site = "https://cricclubs.com/BayerischerCricketVerbandeV/ballbyball.do?matchId=1300&clubId=40958"

# initializing chrome options for chrome driver
chrome_options = webdriver.ChromeOptions()
chrome_options.add_argument("--start-maximized")
chrome_options.add_argument("--headless")

# initializing chrome driver
driver = webdriver.Chrome(options=chrome_options)

# opening the site
driver.get(site)

driver.find_element(By.ID, "ballByBallTeamTab1").click()
time.sleep(1)

# creating soup and getting data like league name team name and date
soup = BeautifulSoup(driver.page_source, "html.parser")
h3_in_mainDiv = soup.select_one("#mainDiv .container h3" ).text
h3_in_mainDiv = h3_in_mainDiv.split('\n')
league = h3_in_mainDiv[0]
match_date = h3_in_mainDiv[-1].strip()
team1_name = soup.select_one("#ballByBallTeamTab1").text
team2_name = soup.select_one("#ballByBallTeamTab2").text

# creating xlsx file
file_name = f"{team1_name} vs {team2_name}.xlsx"

columns_for_df = ["League","Date","Innings", "Ball", "Batting Team", "Bowling Team", "Striker", "Bowler", "Run off bat", "Extras",
                  "Wides", "No Balls", "Byes", "Leg Byes", "penalty", "Wicket Type", "Player Dismissed",]
df_emp = pd.DataFrame(columns=columns_for_df, index=[0])
df_emp.to_excel(file_name)


def parse_and_save_bbd(data, batting_team, bowling_team, innings):
    for index, n in enumerate(data):
        ball_tag = n.select_one(".col2 .ov")
        if ball_tag and len(ball_tag.text.strip()) != 0:
            # initializing variables
            ball = ""
            bowler = ""
            batsman = ""
            wicket_type = ""
            player_dismissed = ""
            wide = ""
            penalty = ""
            leg_bye = ""
            bye = ""
            no_ball = ""
            extra = ""
            run_off_bat = "0"

            # taking ball data
            ball = ball_tag.text.strip()

            # getting all the data in col3(about bowler, batsman, runs, wide etc)
            other_match_data = n.select_one(".col3").text.strip()

            if other_match_data.find("RETIRED") != -1:
                print(other_match_data.split("to"),other_match_data)
            # seperating bowler name from other data
            bowler, other_things, *useless = other_match_data.split("to")

            # extracting batsman from other_things
            if other_things.find("OUT!") != -1:

                out_data = other_things.split("\n")
                if len(out_data) != 1:
                    batsman = out_data[0].split("OUT!")[0]
                    player_dismissed = out_data[3]
                    wicket_type = out_data[1]

                else:
                    batsman = out_data[0].split("OUT!")[0]
                    player_dismissed = batsman
                    wicket_type = "OUT"
            elif other_things.find("WIDE") != -1:
                # if someone scored in wide
                if other_things[-1] == "S":
                    end = other_things.find("WIDE")
                    start = end - 3
                    wide = other_things[start: end]
                else:
                    wide = 1
                batsman = other_things.split("WIDE")[0]
                extra = wide
            else:
                if other_things.find("NO BALL") != -1:
                    no_ball = 1
                    batsman = other_things.split(",")[0]
                    if other_things.find("LEG BYE") != -1:
                        # if it is no ball with leg bye
                        if other_things[-1] == "S":
                            leg_bye_loc = other_things.find("LEG BYE")
                            leg_bye = other_things[leg_bye_loc-3:leg_bye_loc]
                        else:
                            leg_bye = 1
                    elif other_things.find("BYE") != -1:
                        # if it is no ball with bye
                        if other_things[-1] == "S":
                            bye_loc = other_things.find("BYE")
                            bye = other_things[bye_loc-3:bye_loc]
                        else:
                            bye = 1
                    
                    # count extra
                    if leg_bye:
                        extra = int(leg_bye) + no_ball
                    elif bye:
                        extra = int(bye) + no_ball

                elif other_things.find("LEG BYE") != -1:
                    batsman = other_things.split(",")[0]
                    comma_loc = other_things.find(",")
                    run_loc = other_things.find("run")
                    run_off_leg_bye = other_things[comma_loc+1:run_loc]
                    leg_bye = extra = run_off_leg_bye

                elif other_things.find("BYE") != -1:
                    batsman = other_things.split(",")[0]
                    comma_loc = other_things.find(",")
                    run_loc = other_things.find("run")
                    run_off_bye = other_things[comma_loc+1:run_loc]
                    bye = extra = run_off_bye

                else:
                    batsman = other_things.split(",")[0]
                    comma_loc = other_things.find(",")
                    run_loc = other_things.find("run")
                    run_off_bat = other_things[comma_loc+1:run_loc]

            bbd_current = {"League":league,"Date":match_date,"Innings":innings,"Ball": ball,"Batting Team": batting_team,
                           "Bowling Team":bowling_team, "Striker": batsman,"Bowler": bowler, 
                           "Run off bat": run_off_bat, "Extras": extra,"Wides":wide,"No Balls":no_ball,
                           "Byes":bye,"Leg Byes": leg_bye,  "penalty":penalty,"Wicket Type": wicket_type,
                           "Player Dismissed":player_dismissed}
            df_bbd = pd.DataFrame(
                bbd_current, columns=columns_for_df, index=[0])
            with pd.ExcelWriter(file_name, mode="a", if_sheet_exists="overlay", engine="openpyxl") as writer:
                df_bbd.to_excel(writer, startrow=writer.sheets["Sheet1"].max_row, index=[0], header=False)


# getting and saving 1st team's data
bb_html = soup.select('.summary-list .active .ball-by-ball-section .bbb-row')
parse_and_save_bbd(bb_html,team1_name,team2_name,1)

# getting ans saving 2nd team's data
driver.find_element(By.ID,"ballByBallTeamTab2").click()
time.sleep(1)
soup = BeautifulSoup(driver.page_source, "html.parser")
bb_html = soup.select('.summary-list .active .ball-by-ball-section .bbb-row')
parse_and_save_bbd(bb_html,team2_name,team1_name,2)
print("Success")