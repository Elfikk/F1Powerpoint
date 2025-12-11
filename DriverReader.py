import constants2025 as con

class DCReader2025():
    def __init__(self, sheet):
        self.sheet = sheet

    def gather_data(self, number_of_players = con.NUMBER_OF_PLAYERS, number_of_drivers = con.NUMBER_OF_DRIVERS):
        drivers = [self.sheet.cell(i + 2, column = 1).internal_value for i in range(1, number_of_drivers + 1)]
        actual_order = [self.sheet.cell(i + 2, column = 2).internal_value for i in range(1, number_of_drivers + 1)]
        player_orders = {self.sheet.cell(2, column = j+2).internal_value:
            {self.sheet.cell(i+2, column = 1).internal_value:
             (self.sheet.cell(i + 2, column = j+2).internal_value,
              self.sheet.cell(i + 2, column = j+2+number_of_players).internal_value)
              for i in range(1, number_of_drivers + 1)}
            for j in range(1, number_of_players + 1)}
        player_scores = {self.sheet.cell(2, column = j).internal_value:
            self.sheet.cell(25, column = j).internal_value for j in range(10, 10+number_of_players)}

        self.driver_real_order = ["" for i in range(number_of_drivers)]
        for i, order in enumerate(actual_order):
            self.driver_real_order[order-1] = drivers[i]
        # self.player_orders = {player: {drivers[j]: player_orders[player][j] for j in range(number_of_drivers)} for player in player_orders}
        self.player_orders = player_orders
        self.player_scores = player_scores

    def format_to_slide(self):
        return (self.driver_real_order, self.player_orders, self.player_scores)

    def get_scores(self):
        return self.player_scores