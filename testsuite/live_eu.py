from scripts.performance_testing.output import performance_output


class AmsinNonEu(performance_output.AmsinNonEuOutput):
    def __init__(self):
        self.eu = 'yes'
        self.login_server = 'ams'
        super(AmsinNonEu, self).__init__()

    def api_response_time(self):
        self.get_tenant_details('pyappe1')
        self.get_all_entity_properties('pyappe1')
        self.group_by_catalog_masters('pyappe1')
        self.get_all_candidates('pyappe1')
        self.getTestUsersForTest('pyappe1')
        self.interviews('pyappe1')
        self.new_interviews('pyappe1')


Object = AmsinNonEu()
Object.common_login("live_eu")
if Object.login == 'OK':
    Object.api_response_time()
    Object.create_pandas_excel('LIVE_EU')
