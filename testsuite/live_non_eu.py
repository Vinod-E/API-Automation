from scripts.performance_testing.output import performance_output


class AmsinNonEu(performance_output.AmsinNonEuOutput):
    def __init__(self):
        self.eu = 'no'
        self.login_server = 'ams'
        super(AmsinNonEu, self).__init__()

    def api_response_time(self):
        self.get_tenant_details('crpo')
        self.get_all_entity_properties('crpo')
        self.group_by_catalog_masters('crpo')
        self.get_all_candidates('crpo')
        self.getTestUsersForTest('crpo')
        self.interviews('crpo')
        self.new_interviews('crpo')


Object = AmsinNonEu()
Object.common_login("live_non_eu")
if Object.login == 'OK':
    Object.api_response_time()
    Object.create_pandas_excel('LIVE_NON_EU')
