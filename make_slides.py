import openpyxl as px
from pptx import Presentation

from pathlib import Path

from BooleanSlide import BooleanSlide
import constants as con
from ChampionshipReader import ChampionshipReader2025
from ChampionshipSlides import ChampionshipSlides
from make_ppt import (
    make_title_layout,
    make_title_box,
    make_player_layout,
    set_player_detail
)
import question_descriptions as qd
import slideInfoGathering as sig
import SlideTypes as st

def update_score(running_score, new_scores):
    return {player: int(running_score[player]) + int(new_scores[player]) for player in running_score}

def main():

    spreadsheet_path = Path("C:\Projekty\Coding\Python\F1PredsPPT\F12025 Predictions Tracking.xlsx")
    wb = px.open(spreadsheet_path, data_only=True)

    prs = Presentation("examples/ThemeExample.pptx")
    prs.slide_width = 12192 * 1000

    running_scores = {player: "0" for player in con.PLAYERS_TO_COLOURS}

    ## H2H

    # Quali

    intro = st.IntroSlide(
        prs,
        "Quali H2H",
        qd.descriptions["QualiH2H"],
        running_scores
    )
    intro.make_slide()

    quali_h2h_reader = sig.H2HReader(wb["QualiH2H"])
    quali_h2h_reader.gather_data()
    quali_h2h_data = quali_h2h_reader.format_to_slide()
    quali_h2h_partial_scores = quali_h2h_reader.get_scores()

    for i in range(len(quali_h2h_data)):
        data = quali_h2h_data[i]
        h2h_eg = st.H2HSlide(prs, *data)
        h2h_eg.make_slide()

        driver1, driver2 = data[0]
        title_text = f"Quali H2H: {driver1} VS {driver2}"
        title_shape = make_title_layout(prs, h2h_eg.slide)
        make_title_box(title_text, title_shape)

        player_footer = make_player_layout(prs, h2h_eg.slide, len(data[2]), 4/25)
        player_deets = [(player, quali_h2h_partial_scores[player][i]) for player in quali_h2h_partial_scores]
        set_player_detail(player_deets, player_footer)

    new_scores = quali_h2h_reader.get_final_scores()
    running_scores = update_score(running_scores, new_scores)

    quali_h2h_roundup = st.RoundUpSlide(prs, running_scores, "Quali H2H")

    # Race

    intro = st.IntroSlide(
        prs,
        "Race H2H",
        qd.descriptions["RaceH2H"],
        running_scores
    )
    intro.make_slide()

    race_h2h_reader = sig.H2HReader(wb["RaceH2H"])
    race_h2h_reader.gather_data()
    race_h2h_data = race_h2h_reader.format_to_slide()
    race_h2h_partial_scores = race_h2h_reader.get_scores()

    for i in range(len(race_h2h_data)):
        data = race_h2h_data[i]

        h2h_eg = st.H2HSlide(prs, *data)
        h2h_eg.make_slide()

        driver1, driver2 = data[0]
        title_text = f"Race H2H: {driver1} VS {driver2}"
        title_shape = make_title_layout(prs, h2h_eg.slide)
        make_title_box(title_text, title_shape)

        player_footer = make_player_layout(prs, h2h_eg.slide, len(data[2]), 4/25)
        player_deets = [(player, race_h2h_partial_scores[player][i] + quali_h2h_partial_scores[player][-1]) for player in race_h2h_partial_scores]
        set_player_detail(player_deets, player_footer)

    new_scores = race_h2h_reader.get_final_scores()

    running_scores = update_score(running_scores, new_scores)
    race_h2h_roundup = st.RoundUpSlide(prs, running_scores, "Race H2H")

    ## Constructor Predictions

    intro = st.IntroSlide(prs,
                          "Constructor's Championship",
                          qd.descriptions["ConstructorPredictions"],
                          running_scores)
    intro.make_slide()

    cc_reader = ChampionshipReader2025(wb["ConstructorPredictions"])
    cc_reader.gather_data(number_of_competitors=10)
    cc_data = cc_reader.format_to_slide()

    drivers_slide = ChampionshipSlides(
        prs,
        cc_data[0],
        cc_data[1],
        cc_data[2],
        "Constructor's Championship")

    drivers_slide.make_slide()

    new_scores = cc_reader.get_scores()
    print(new_scores)
    running_scores = update_score(running_scores, new_scores)
    cc_roundup = st.RoundUpSlide(prs, running_scores, "Constructor's Championship")

    ## Driver Predictions

    intro = st.IntroSlide(prs,
                          "Driver's Championship",
                          qd.descriptions["DriverPredictions"],
                          running_scores)
    intro.make_slide()

    dc_reader = ChampionshipReader2025(wb["DriverPredictions"])
    dc_reader.gather_data(number_of_competitors=20)
    dc_data = dc_reader.format_to_slide()

    drivers_slide = ChampionshipSlides(
        prs,
        dc_data[0],
        dc_data[1],
        dc_data[2],
        "Driver's Championship")

    drivers_slide.make_slide()

    new_scores = dc_reader.get_scores()
    running_scores = update_score(running_scores, new_scores)
    dc_roundup = st.RoundUpSlide(prs, running_scores, "Driver's Championship")

    ## True False

    intro = st.IntroSlide(prs,
                          "True/False",
                          qd.descriptions["True/False"],
                          running_scores)
    intro.make_slide()

    # Q2 Eliminations
    title_text = f"True/False: 5+ Q2 Eliminations"
    intro = st.IntroSlide(prs,
                          title_text,
                          qd.descriptions["Q2Elims"],
                          running_scores)
    intro.make_slide()

    q2_reader = sig.TrueFalseReader(wb["Q2Elims"])
    q2_reader.gather_data()
    q2_data = q2_reader.format_to_slide()

    print(q2_data[0])
    print(q2_data[1])

    q2_slides = BooleanSlide(prs, *q2_data, title_text)
    q2_slides.make_slide()

    new_scores = q2_reader.get_scores()
    running_scores = update_score(running_scores, new_scores)
    q2_roundup = st.RoundUpSlide(prs, running_scores, title_text)

    # Pole Positions

    title_text = f"True/False: Pole Positions"

    intro = st.IntroSlide(prs,
                          title_text,
                          qd.descriptions["PolePositions"],
                          running_scores)
    intro.make_slide()

    pole_reader = sig.TrueFalseReader(wb["PolePositions"])
    pole_reader.gather_data()
    pole_data = pole_reader.format_to_slide()

    pole_slides = BooleanSlide(prs, *pole_data, title_text)
    pole_slides.make_slide()

    new_scores = pole_reader.get_scores()
    running_scores = update_score(running_scores, new_scores)
    pole_roundup = st.RoundUpSlide(prs, running_scores, title_text)

    # Lap1DNF

    title_text = f"True/False: Lap 1 DNFs"

    intro = st.IntroSlide(prs,
                          title_text,
                          qd.descriptions["Lap1DNF"],
                          running_scores)
    intro.make_slide()

    pens1dnf_reader = sig.TrueFalseReader(wb["Lap1DNF"])
    pens1dnf_reader.gather_data()
    pens1dnf_data = pens1dnf_reader.format_to_slide()

    pens1dnf_slides = BooleanSlide(prs, *pens1dnf_data, title_text)
    pens1dnf_slides.make_slide()

    new_scores = pens1dnf_reader.get_scores()
    print(running_scores)
    print(new_scores)
    running_scores = update_score(running_scores, new_scores)
    pens1dnf_roundup = st.RoundUpSlide(prs, running_scores, "Lap 1 DNFs")

    # Wins

    title_text = f"True/False: Wins"
    intro = st.IntroSlide(prs,
                          title_text,
                          qd.descriptions["Wins"],
                          running_scores)
    intro.make_slide()

    win_reader = sig.TrueFalseReader(wb["Wins"])
    win_reader.gather_data()
    win_data = win_reader.format_to_slide()

    win_slides = BooleanSlide(prs, *win_data, title_text)
    win_slides.make_slide()

    new_scores = win_reader.get_scores()
    running_scores = update_score(running_scores, new_scores)
    win_roundup = st.RoundUpSlide(prs, running_scores, title_text)

    # DOTD
    title_text = f"True/False: Driver of the Day"
    intro = st.IntroSlide(prs,
                          title_text,
                          qd.descriptions["DOTD"],
                          running_scores)
    intro.make_slide()

    dotd_reader = sig.TrueFalseReader(
        wb["DOTD"],
        shift = 0
        )
    dotd_reader.gather_data()
    dotd_data = dotd_reader.format_to_slide()

    dotd_slides = BooleanSlide(prs, *dotd_data, title_text)
    dotd_slides.make_slide()

    new_scores = dotd_reader.get_scores()
    running_scores = update_score(running_scores, new_scores)
    dotd_roundup = st.RoundUpSlide(prs, running_scores, title_text)

    ## Superlatives

    intro = st.IntroSlide(prs,
                          "Superlatives",
                          qd.descriptions["Superlatives"],
                          running_scores)
    intro.make_slide()

    prs.save("konkurenz2025.pptx")

    # DNFs
    title_text = f"Superlatives: DNFs"
    intro = st.IntroSlide(prs,
                          title_text,
                          qd.descriptions["DNFs"],
                          running_scores)
    intro.make_slide()

    dnf_reader = sig.SuperlativeDriverReader(wb["DNFs"], player_col = 29)
    dnf_reader.gather_data()
    dnf_data = dnf_reader.format_to_slide()

    dnf_slide = st.SuperlativeSlide(prs, *dnf_data)
    dnf_slide.make_slide()

    title_shape = make_title_layout(prs, dnf_slide.slide, 4/25)
    make_title_box(title_text, title_shape)

    new_scores = dnf_reader.get_scores()
    running_scores = update_score(running_scores, new_scores)
    dnf_roundup = st.RoundUpSlide(prs, running_scores, title_text)

    # Pit Stops

    title_text = f"Superlatives: Most Pit Stops"
    intro = st.IntroSlide(prs,
                          title_text,
                          qd.descriptions["PitStops"],
                          running_scores)
    intro.make_slide()

    pit_reader = sig.SuperlativeDriverReader(wb["PitStops"], player_col = 29)
    pit_reader.gather_data()
    pit_data = pit_reader.format_to_slide()

    pit_slide = st.SuperlativeSlide(prs, *pit_data)
    pit_slide.make_slide()

    title_shape = make_title_layout(prs, pit_slide.slide, 4/25)
    make_title_box(title_text, title_shape)

    new_scores = pit_reader.get_scores()
    running_scores = update_score(running_scores, new_scores)
    pit_roundup = st.RoundUpSlide(prs, running_scores, title_text)

    # Pens

    title_text = f"Superlatives: Penalties"
    intro = st.IntroSlide(prs,
                          title_text,
                          qd.descriptions["Pens"],
                          running_scores)
    intro.make_slide()

    pens_reader = sig.SuperlativeDriverReader(
        wb["Pens"],
        player_col=29
        )
    pens_reader.gather_data()
    pens_data = pens_reader.format_to_slide()

    pens_slide = st.SuperlativeSlide(prs, *pens_data)
    pens_slide.make_slide()

    title_shape = make_title_layout(prs, pens_slide.slide, 4/25)
    make_title_box(title_text, title_shape)

    new_scores = pens_reader.get_scores()
    running_scores = update_score(running_scores, new_scores)
    pens_roundup = st.RoundUpSlide(prs, running_scores, title_text)

    # Points Improver

    title_text = f"Superlatives: Largest Point Improver Between Season Halves"
    intro = st.IntroSlide(prs,
                          title_text,
                          qd.descriptions["SecondWind"],
                          running_scores)
    intro.make_slide()

    delta_reader = sig.SuperlativeDriverReader(
        wb["SecondWind"],
        metric_col=5,
        rank_col=6,
        player_col=9
    )
    delta_reader.gather_data()
    delta_data = delta_reader.format_to_slide()

    delta_slide = st.SuperlativeSlide(prs, *delta_data)
    delta_slide.make_slide()

    title_shape = make_title_layout(prs, delta_slide.slide, 4/25)
    make_title_box(title_text, title_shape)

    new_scores = delta_reader.get_scores()
    running_scores = update_score(running_scores, new_scores)
    delta_roundup = st.RoundUpSlide(prs, running_scores, title_text)

    # Slow Starter

    title_text = f"Superlatives: Most Races to Score"
    intro = st.IntroSlide(prs,
                          title_text,
                          qd.descriptions["SlowStarter"],
                          running_scores)
    intro.make_slide()

    ss_reader = sig.SuperlativeDriverReader(
        wb["SlowStarter"],
        metric_col=3,
        rank_col=4,
        player_col=30
        )
    ss_reader.gather_data()
    ss_data = ss_reader.format_to_slide()

    ss_slide = st.SuperlativeSlide(prs, *ss_data)
    ss_slide.make_slide()

    title_shape = make_title_layout(prs, ss_slide.slide, 4/25)
    make_title_box(title_text, title_shape)

    new_scores = ss_reader.get_scores()
    running_scores = update_score(running_scores, new_scores)
    ss_roundup = st.RoundUpSlide(prs, running_scores, title_text)

    ## Superlative Teams

    # Engine Components
    title_text = f"Superlatives: Most Engine Components Used By Teams"
    intro = st.IntroSlide(prs,
                          title_text,
                          qd.descriptions["Ngin"],
                          running_scores)
    intro.make_slide()

    ngin_reader = sig.SuperlativeTeamReader(wb["Ngin"])
    ngin_reader.gather_data()
    ngin_data = ngin_reader.format_to_slide()

    ngin_slide = st.SuperlativeSlide(prs, *ngin_data, "Team")
    ngin_slide.make_slide()

    title_shape = make_title_layout(prs, ngin_slide.slide, 4/25)
    make_title_box(title_text, title_shape)

    new_scores = ngin_reader.get_scores()
    running_scores = update_score(running_scores, new_scores)
    ngin_roundup = st.RoundUpSlide(prs, running_scores, title_text)

    # Quali Constructor
    title_text = f"Superlatives: Lowest Qualifying Average by Team"
    intro = st.IntroSlide(prs,
                          title_text,
                          qd.descriptions["QualiConstructor"],
                          running_scores)
    intro.make_slide()

    qc_reader = sig.SuperlativeTeamReader(
        wb["QualiConstructor"],
        team_col = 32,
        metric_col = 35,
        rank_col = 34,
        player_col = 32,
        prediction_col = 33,
        score_col = 35,
        team_row = 4,
        player_row = 16,
        team_row_delta=1
    )
    qc_reader.gather_data()
    qc_data = qc_reader.format_to_slide()

    qc_slide = st.SuperlativeSlide(prs, *qc_data, "Team")
    qc_slide.make_slide()

    title_shape = make_title_layout(prs, qc_slide.slide, 4/25)
    make_title_box(title_text, title_shape)

    new_scores = qc_reader.get_scores()
    running_scores = update_score(running_scores, new_scores)
    qc_roundup = st.RoundUpSlide(prs, running_scores, title_text)

    # Constructor PSs

    title_text = f"Superlatives: Closest Finishing Position Team"
    intro = st.IntroSlide(prs,
                          title_text,
                          qd.descriptions["ClosestTeam"],
                          running_scores)
    intro.make_slide()

    closest_team_reader = sig.SuperlativeTeamReader(
        wb["ClosestTeam"],
        team_col = 32,
        metric_col = 35,
        rank_col = 34,
        player_col = 32,
        prediction_col = 33,
        score_col = 35,
        team_row = 4,
        player_row = 16,
        team_row_delta=1
    )
    closest_team_reader.gather_data()
    closest_team_data = closest_team_reader.format_to_slide()

    closest_team_slide = st.SuperlativeSlide(prs, *closest_team_data, "Team")
    closest_team_slide.make_slide()

    title_shape = make_title_layout(prs, closest_team_slide.slide, 4/25)
    make_title_box(title_text, title_shape)

    new_scores = closest_team_reader.get_scores()
    running_scores = update_score(running_scores, new_scores)
    closest_team_roundup = st.RoundUpSlide(prs, running_scores, title_text)

    ## Pick 5
    title_text = "Pick 5 Races"
    intro = st.IntroSlide(prs,
                          title_text,
                          qd.descriptions["Pick5Races"],
                          running_scores)
    intro.make_slide()

    # Yuki
    title_text = "Pick 5 Races - Yuki Quali"
    intro = st.IntroSlide(prs,
                          title_text,
                          qd.descriptions["PastnPukious"],
                          running_scores)
    intro.make_slide()

    yuki_reader = sig.Pick5RacesReader(wb["PastnPukious"])
    yuki_reader.gather_data()
    yuki_data = yuki_reader.format_to_slide()

    yuki_slides = st.Pick5RacesSlide(prs, *yuki_data, running_scores.copy(), "Pick 5 Races - Yuki")
    yuki_slides.make_slide()

    new_scores = yuki_reader.get_scores()
    running_scores = update_score(running_scores, new_scores)
    yuki_roundup = st.RoundUpSlide(prs, running_scores, title_text)

    # Bearman
    title_text = "Pick 5 Races - Bearman Races"
    intro = st.IntroSlide(prs,
                          title_text,
                          qd.descriptions["BearHug"],
                          running_scores)
    intro.make_slide()

    bear_reader = sig.Pick5RacesReader(wb["BearHug"])
    bear_reader.gather_data()
    bear_data = bear_reader.format_to_slide()

    bear_slides = st.Pick5RacesSlide(prs, *bear_data, running_scores.copy(), "Pick 5 Races - Bearman")
    bear_slides.make_slide()

    new_scores = bear_reader.get_scores()
    running_scores = update_score(running_scores, new_scores)
    bear_roundup = st.RoundUpSlide(prs, running_scores, title_text)

    # Blowy Engines
    title_text = "Pick 5 Races - Blowy Engines"
    intro = st.IntroSlide(prs,
                          title_text,
                          qd.descriptions["BlowyNgin"],
                          running_scores)
    intro.make_slide()

    blowy_reader = sig.Pick5RacesReader(wb["BlowyNgin"])
    blowy_reader.gather_data()
    blowy_data = blowy_reader.format_to_slide()

    blowy_slides = st.Pick5RacesSlide(prs, *blowy_data, running_scores.copy(), "Pick 5 Races - Blowy Engines")
    blowy_slides.make_slide()

    new_scores = blowy_reader.get_scores()
    running_scores = update_score(running_scores, new_scores)
    blowy_roundup = st.RoundUpSlide(prs, running_scores, title_text)

    prs.save("konkurenz2025.pptx")

main()
