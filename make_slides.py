import openpyxl as px
from pptx import Presentation

from pathlib import Path

import constants2024 as con
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

    spreadsheet_path = Path("C:\Projekty\Coding\Python\F1PredsPPT\F12024 Predictions Tracking.xlsx")
    wb = px.open(spreadsheet_path, data_only=True)

    prs = Presentation()
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

    cc_reader = sig.CCReader(wb["ConstructorPredictions"])
    cc_reader.gather_data()
    cc_data = cc_reader.format_to_slide()

    cc_slides = st.CCSlides(prs, *cc_data)
    cc_slides.make_slide()

    new_scores = cc_reader.get_scores()
    running_scores = update_score(running_scores, new_scores)
    cc_roundup = st.RoundUpSlide(prs, running_scores, "Constructor's Championship")

    ## Driver Predictions

    intro = st.IntroSlide(prs,
                          "Driver's Championship",
                          qd.descriptions["DriverPredictions"],
                          running_scores)
    intro.make_slide()

    dc_reader = sig.DCReader(wb["DriverPredictions"])
    dc_reader.gather_data()
    dc_data = dc_reader.format_to_slide()

    dc_slides = st.DCSlides(prs, *dc_data)
    dc_slides.make_slide()

    new_scores = dc_reader.get_scores()
    running_scores = update_score(running_scores, new_scores)
    dc_roundup = st.RoundUpSlide(prs, running_scores, "Driver's Championship")

    ## First 6 Races

    intro = st.IntroSlide(prs,
                          "N After N",
                          qd.descriptions["First6Races"],
                          running_scores)
    intro.make_slide()

    first6_reader = sig.First6RacesReader(wb["First6Races"])
    first6_reader.gather_data()
    first6_data = first6_reader.format_to_slide()

    first6_slides = st.First6RacesSlides(prs, *first6_data)
    first6_slides.make_slide()

    new_scores = first6_reader.get_scores()
    running_scores = update_score(running_scores, new_scores)
    first6_roundup = st.RoundUpSlide(prs, running_scores, "N After N")

    ## True False

    intro = st.IntroSlide(prs,
                          "True/False",
                          qd.descriptions["True/False"],
                          running_scores)
    intro.make_slide()

    # Podiums

    title_text = f"True/False: Podiums"

    intro = st.IntroSlide(prs,
                          title_text,
                          qd.descriptions["Podiums"],
                          running_scores)
    intro.make_slide()

    podium_reader = sig.TrueFalseReader(wb["Podiums"])
    podium_reader.gather_data()
    podium_data = podium_reader.format_to_slide()

    podium_slides = st.TrueFalseSlide(prs, *podium_data, False)
    podium_slides.make_slide()

    title_shape = make_title_layout(prs, podium_slides.slide, 4/25)
    make_title_box(title_text, title_shape)

    podium_slides_score = st.TrueFalseSlide(prs, *podium_data, True)
    podium_slides_score.make_slide()

    title_shape = make_title_layout(prs, podium_slides_score.slide, 4/25)
    make_title_box(title_text, title_shape)

    new_scores = podium_reader.get_scores()
    print(running_scores)
    print(new_scores)
    running_scores = update_score(running_scores, new_scores)
    podium_roundup = st.RoundUpSlide(prs, running_scores, "Podiums")

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

    pole_slides = st.TrueFalseSlide(prs, *pole_data, False)
    pole_slides.make_slide()

    title_shape = make_title_layout(prs, pole_slides.slide, 4/25)
    make_title_box(title_text, title_shape)

    pole_slides_score = st.TrueFalseSlide(prs, *pole_data, True)
    pole_slides_score.make_slide()

    title_shape = make_title_layout(prs, pole_slides_score.slide, 4/25)
    make_title_box(title_text, title_shape)

    new_scores = pole_reader.get_scores()
    running_scores = update_score(running_scores, new_scores)
    pole_roundup = st.RoundUpSlide(prs, running_scores, title_text)

    # FLs

    title_text = f"True/False: Fastest Laps"
    intro = st.IntroSlide(prs,
                          title_text,
                          qd.descriptions["FLs"],
                          running_scores)
    intro.make_slide()

    fl_reader = sig.TrueFalseReader(wb["FLs"])
    fl_reader.gather_data()
    fl_data = fl_reader.format_to_slide()

    fl_slides = st.TrueFalseSlide(prs, *fl_data, False)
    fl_slides.make_slide()

    title_shape = make_title_layout(prs, fl_slides.slide, 4/25)
    make_title_box(title_text, title_shape)

    fl_slides_score = st.TrueFalseSlide(prs, *fl_data, True)
    fl_slides_score.make_slide()

    title_shape = make_title_layout(prs, fl_slides_score.slide, 4/25)
    make_title_box(title_text, title_shape)

    new_scores = fl_reader.get_scores()
    running_scores = update_score(running_scores, new_scores)
    fl_roundup = st.RoundUpSlide(prs, running_scores, title_text)

    # Q1 Eliminations
    title_text = f"True/False: 5+ Q1 Eliminations"
    intro = st.IntroSlide(prs,
                          title_text,
                          qd.descriptions["Q1Elims"],
                          running_scores)
    intro.make_slide()

    q1_reader = sig.TrueFalseReader(wb["Q1Elims"])
    q1_reader.gather_data()
    q1_data = q1_reader.format_to_slide()

    q1_slides = st.TrueFalseSlide(prs, *q1_data, False)
    q1_slides.make_slide()

    title_shape = make_title_layout(prs, q1_slides.slide, 4/25)
    make_title_box(title_text, title_shape)

    q1_slides_score = st.TrueFalseSlide(prs, *q1_data, True)
    q1_slides_score.make_slide()

    title_shape = make_title_layout(prs, q1_slides_score.slide, 4/25)
    make_title_box(title_text, title_shape)

    new_scores = q1_reader.get_scores()
    running_scores = update_score(running_scores, new_scores)
    q1_roundup = st.RoundUpSlide(prs, running_scores, title_text)

    # Monaco
    title_text = f"True/False: Monaco Top 10"
    intro = st.IntroSlide(prs,
                          title_text,
                          qd.descriptions["Monaco"],
                          running_scores)
    intro.make_slide()

    monaco_reader = sig.TrueFalseReader(
        wb["Monaco"],
        shift = 0
        )
    monaco_reader.gather_data()
    monaco_data = monaco_reader.format_to_slide()

    monaco_slides = st.TrueFalseSlide(prs, *monaco_data, False)
    monaco_slides.make_slide()

    # title_text = f"True/False: Monaco Top 10"
    title_shape = make_title_layout(prs, monaco_slides.slide, 4/25)
    make_title_box(title_text, title_shape)

    monaco_slides_score = st.TrueFalseSlide(prs, *monaco_data, True)
    monaco_slides_score.make_slide()

    title_shape = make_title_layout(prs, monaco_slides_score.slide, 4/25)
    make_title_box(title_text, title_shape)

    new_scores = monaco_reader.get_scores()
    running_scores = update_score(running_scores, new_scores)
    monaco_roundup = st.RoundUpSlide(prs, running_scores, title_text)

    ## Superlatives

    intro = st.IntroSlide(prs,
                          "Superlatives",
                          qd.descriptions["Superlatives"],
                          running_scores)
    intro.make_slide()

    # DNFs
    title_text = f"Superlatives: DNFs"
    intro = st.IntroSlide(prs,
                          title_text,
                          qd.descriptions["DNFs"],
                          running_scores)
    intro.make_slide()

    dnf_reader = sig.SuperlativeDriverReader(wb["DNFs"])
    dnf_reader.gather_data()
    dnf_data = dnf_reader.format_to_slide()

    dnf_slide = st.SuperlativeSlide(prs, *dnf_data)
    dnf_slide.make_slide()

    title_shape = make_title_layout(prs, dnf_slide.slide, 4/25)
    make_title_box(title_text, title_shape)

    new_scores = dnf_reader.get_scores()
    running_scores = update_score(running_scores, new_scores)
    dnf_roundup = st.RoundUpSlide(prs, running_scores, title_text)

    # Laps

    title_text = f"Superlatives: Lap %"
    intro = st.IntroSlide(prs,
                          title_text,
                          qd.descriptions["Laps"],
                          running_scores)
    intro.make_slide()

    lap_reader = sig.SuperlativeDriverReader(
        wb["Laps"],
        metric_col=4,
        rank_col=5,
        player_col=32
        )
    lap_reader.gather_data()
    lap_data = lap_reader.format_to_slide()

    lap_slide = st.SuperlativeSlide(prs, *lap_data)
    lap_slide.make_slide()

    title_shape = make_title_layout(prs, lap_slide.slide, 4/25)
    make_title_box(title_text, title_shape)

    new_scores = lap_reader.get_scores()
    running_scores = update_score(running_scores, new_scores)
    lap_roundup = st.RoundUpSlide(prs, running_scores, title_text)

    # Pit Stops

    title_text = f"Superlatives: Most Pit Stops"
    intro = st.IntroSlide(prs,
                          title_text,
                          qd.descriptions["PitStops"],
                          running_scores)
    intro.make_slide()

    pit_reader = sig.SuperlativeDriverReader(wb["PitStops"])
    pit_reader.gather_data()
    pit_data = pit_reader.format_to_slide()

    pit_slide = st.SuperlativeSlide(prs, *pit_data)
    pit_slide.make_slide()

    title_shape = make_title_layout(prs, pit_slide.slide, 4/25)
    make_title_box(title_text, title_shape)

    new_scores = pit_reader.get_scores()
    running_scores = update_score(running_scores, new_scores)
    pit_roundup = st.RoundUpSlide(prs, running_scores, title_text)

    # Slow Starter

    title_text = f"Superlatives: Most Races to Score"
    intro = st.IntroSlide(prs,
                          title_text,
                          qd.descriptions["SlowStarter"],
                          running_scores)
    intro.make_slide()

    ss_reader = sig.SuperlativeDriverReader(
        wb["SlowStarter"],
        rank_col=4,
        player_col=31
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

    # Points Improver

    title_text = f"Superlatives: Largest Point Improver 2023 to 2024"
    intro = st.IntroSlide(prs,
                          title_text,
                          qd.descriptions["PointsImprover"],
                          running_scores)
    intro.make_slide()

    delta_reader = sig.SuperlativeDriverReader(
        wb["PointsImprover"],
        metric_col=4,
        rank_col=5,
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

    # Position Improver

    title_text = f"Superlatives: Most Positions Gained Per Race"
    intro = st.IntroSlide(prs,
                          title_text,
                          qd.descriptions["PositionImprover"],
                          running_scores)
    intro.make_slide()

    pos_reader = sig.SuperlativeDriverReader(
        wb["PositionImprover"],
        metric_col=4,
        player_col=54
    )
    pos_reader.gather_data()
    pos_data = pos_reader.format_to_slide()

    pos_slide = st.SuperlativeSlide(prs, *pos_data)
    pos_slide.make_slide()

    title_shape = make_title_layout(prs, pos_slide.slide, 4/25)
    make_title_box(title_text, title_shape)

    new_scores = pos_reader.get_scores()
    running_scores = update_score(running_scores, new_scores)
    pos_roundup = st.RoundUpSlide(prs, running_scores, title_text)

    ## Superlative Teams

    # Engine Components
    title_text = f"Superlatives: Most Engine Components Used By Teams"
    intro = st.IntroSlide(prs,
                          title_text,
                          qd.descriptions["EngineComponents"],
                          running_scores)
    intro.make_slide()

    ngin_reader = sig.SuperlativeTeamReader(wb["EngineComponents"])
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
        player_col = 39,
        prediction_col = 40,
        score_col = 42,
        team_row = 4,
        player_row = 4,
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

    title_text = f"Superlatives: Fastest Pit Stop by Team"
    intro = st.IntroSlide(prs,
                          title_text,
                          qd.descriptions["ConstructorPSs"],
                          running_scores)
    intro.make_slide()

    ps_reader = sig.SuperlativeTeamReader(
        wb["ConstructorPSs"],
        team_col=1,
        player_col=32,
        prediction_col=33,
        score_col=35,
        team_row_delta=1,
        team_row=3,
        player_row=4,
        metric_col=2,
        rank_col = 3
    )
    ps_reader.gather_data()
    ps_data = ps_reader.format_to_slide()

    pc_slide = st.SuperlativeSlide(prs, *ps_data, "Team")
    pc_slide.make_slide()

    title_shape = make_title_layout(prs, pc_slide.slide, 4/25)
    make_title_box(title_text, title_shape)

    new_scores = ps_reader.get_scores()
    running_scores = update_score(running_scores, new_scores)
    pc_roundup = st.RoundUpSlide(prs, running_scores, title_text)

    ## Pick 5
    title_text = "Pick 5 Races"
    intro = st.IntroSlide(prs,
                          title_text,
                          qd.descriptions["Pick5Races"],
                          running_scores)
    intro.make_slide()

    # Gasly
    title_text = "Pick 5 Races - Gasly Races"
    intro = st.IntroSlide(prs,
                          title_text,
                          qd.descriptions["GaslyPoints"],
                          running_scores)
    intro.make_slide()

    gasly_reader = sig.Pick5RacesReader(wb["GaslyPoints"])
    gasly_reader.gather_data()
    gasly_data = gasly_reader.format_to_slide()

    gasly_slides = st.Pick5RacesSlide(prs, *gasly_data, running_scores.copy(), "Pick 5 Races - Gasly")
    gasly_slides.make_slide()

    new_scores = gasly_reader.get_scores()
    running_scores = update_score(running_scores, new_scores)
    gasly_roundup = st.RoundUpSlide(prs, running_scores, title_text)

    # Hulk
    title_text = "Pick 5 Races - Hulk Qualifying"
    intro = st.IntroSlide(prs,
                          title_text,
                          qd.descriptions["HulkQuali"],
                          running_scores)
    intro.make_slide()

    hulk_reader = sig.Pick5RacesReader(wb["HulkQuali"])
    hulk_reader.gather_data()
    hulk_data = hulk_reader.format_to_slide()

    hulk_slides = st.Pick5RacesSlide(prs, *hulk_data, running_scores.copy(), "Pick 5 Races - Hulk")
    hulk_slides.make_slide()

    new_scores = hulk_reader.get_scores()
    running_scores = update_score(running_scores, new_scores)
    hulk_roundup = st.RoundUpSlide(prs, running_scores, title_text)

    # Blowy Engines
    title_text = "Pick 5 Races - Blowy Engines"
    intro = st.IntroSlide(prs,
                          title_text,
                          qd.descriptions["BlowyEngines"],
                          running_scores)
    intro.make_slide()

    blowy_reader = sig.Pick5RacesReader(wb["BlowyEngines"])
    blowy_reader.gather_data()
    blowy_data = blowy_reader.format_to_slide()

    blowy_slides = st.Pick5RacesSlide(prs, *blowy_data, running_scores.copy(), "Pick 5 Races - Blowy Engines")
    blowy_slides.make_slide()

    new_scores = blowy_reader.get_scores()
    running_scores = update_score(running_scores, new_scores)
    blowy_roundup = st.RoundUpSlide(prs, running_scores, title_text)

    prs.save("theSlidesTM.pptx")

main()