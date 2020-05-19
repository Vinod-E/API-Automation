from scripts.performance_testing.output import chart_analysis


class Graph(chart_analysis.Chart):
    def __init__(self):
        super(Graph, self).__init__()

    def generate_graph(self):
        self.read_data_from_excel('AMSIN_NON_EU')
        self.read_data_from_excel('AMSIN_EU')
        self.read_data_from_excel('LIVE_NON_EU')
        self.read_data_from_excel('LIVE_EU')
        self.chart_sheets('Dashboard')
        self.merge_2_excels()


Object = Graph()
Object.generate_graph()
