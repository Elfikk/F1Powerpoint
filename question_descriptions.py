descriptions = {
    "QualiH2H": """Predict the driver in each pairing that will finish ahead in qualifying most often. Excludes Sprint Qualifying.\n+5pts per correct pick, 0 in a draw""",
    "RaceH2H": """Predict the driver in each pairing that will finish ahead on the road most often (not necessarilly points wise!) in races.Excludes Sprints.\n+5pts per correct pick, 0 in a draw""",
    "DriverPredictions": """
Predict the Driver's Championship order. This question is punishing to having drivers very far away from their actual positions, but if you're close (1-4 positions) you'll score well still.

How it's scored - the lie:
For each driver, we find the difference between their actual position and your prediction. We square it, times by 0.3. Add up all of these, take them away from 200, round down - that's your score.

How it's scored - the truth:
The Spearman's rank coefficient between your order and the actual order is found.
This is mapped such that a score of 1 gets 200 points and 0.5 gets 0. The rest are negative (this has not happened yet).
""",
    "ConstructorPredictions": """
    Predict the Constructor's Championship order. This question is punishing to having teams very far away from their actual positions, but if you're close (1 or 2 positions) you'll score well still.

How it's scored - the lie:
For each team, we find the difference between their actual position and your prediction. We square it, times by 2.4. Add up all of these, take them away from 200, round down - that's your score.

How it's scored - the truth:
The Spearman's rank coefficient between your order and the actual order is found.
This is mapped such that a score of 1 gets 200 points and 0.5 gets 0. The rest are negative (this has not happened yet)."
""",
    "First6Races": """
N After N. Predict the driver who will 1st in the championship after the 1st round, the driver who will be 2nd after the 2nd round, 3rd after the 3rd round...up to the 6th round.

F1 style points - The best prediction at each stage is worth 25pts, the next is 18pts etc.
So for example, getting the 3rd driver in the 3rd round is 25pts, getting the 2nd or 4th is 18pts, etc
""",
    "True/False": """True/False - A set of questions where I made you predict whether a driver will or will not achieve something (good or bad).

Predict all drivers that you believe will...
a) score at least one podium finish.
b) qualify on pole position.
c) get a fastest lap.
d) be eliminated 5 or more times in Q1.
e) finish in the top 10 at Monaco.

Scoring: Each true prediction that comes true is 5pts. -3pts for missing predictions, -3pts for any true prediction that doesn't come true.
""",
    "Podiums": """Predict all drivers that you believe will score at least one podium finish.

Scoring: Each true prediction that comes true is 5pts. -3pts for missing predictions, -3pts for any true prediction that doesn't come true.
""",
    "PolePositions": """Predict all drivers that you believe will qualify on pole position.

Scoring: Each true prediction that comes true is 5pts. -3pts for missing predictions, -3pts for any true prediction that doesn't come true.
""",
    "FLs": """Predict all drivers that you believe will get a fastest lap.

Scoring: Each true prediction that comes true is 5pts. -3pts for missing predictions, -3pts for any true prediction that doesn't come true.
""",
    "Q1Elims": """Predict all drivers that you believe will be eliminated 5 or more times in Q1.

Scoring: Each true prediction that comes true is 5pts. -3pts for missing predictions, -3pts for any true prediction that doesn't come true.
""",
    "Monaco": """Predict all drivers that you believe will finish in the top 10 at Monaco.

Scoring: Each true prediction that comes true is 5pts. -3pts for missing predictions, -3pts for any true prediction that doesn't come true.
""",
    "Superlatives": """
The superlatives - sets of question about the extremes, for drivers and teams.

F1 style points - the best prediction will get you 25 points, the second best gets you 18,...

Draw handling:
If two drivers/teams are equal in a category, their position is averaged and rounded down.
Example (Completely made up):
Yuki and Zhou both 41 pit stops over the season, the most of everyone - giving them rank 1.5, worth 18 points. The next driver is Lando, with 40 pit stops, with rank 3 - giving 15 points.
""",
    "DNFs": "Name the driver with the most DNFs over the season.",
    "Laps": "Name the driver who will complete the lowest proportion of the laps they could have completed.",
    "PositionImprover": "Name the driver that will gain the most positions against their starting grid position on average.",
    "PointsImprover": "Name the driver that will improve their driver's championship points tally the most, as an absolute difference between their 2023 and 2024 points tallies.",
    "PitStops": "Name the driver that will complete the most pit stops.",
    "SlowStarter": "Name the driver that will take the most races to score.",
    "EngineComponents": "Name the team that will use up the most engine components.",
    "ConstructorPSs": "Name the team that will complete the fastest pit stop of the season.",
    "QualiConstructor": "Name the team that will have the highest qualifying average as a team.",
    "Pick5Races": "Pick 5 Races - name 5 races over the season for different scenarios  (that's a bit broad).",
    "GaslyPoints": "Pick 5 Races. You get Gasly's points in them...",
    "HulkQuali": "Pick 5 Races. You get the points Hulkenberg would have gotten if the weekend ended after quali.",
    "BlowyEngines": "Pick 5 Races. You get 5 points for every DNF in that race."
}