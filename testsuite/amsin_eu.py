from scripts.performance_testing.output import performance_output


class AmsinNonEu(performance_output.AmsinNonEuOutput):
    def __init__(self):
        self.eu = 'yes'
        self.login_server = 'amsin'
        super(AmsinNonEu, self).__init__()

    def api_response_time(self):
        self.get_tenant_details('amsin_eu_get_tenant_details')
        self.get_all_entity_properties('amsin_eu_get_all_entity_properties')
        self.group_by_catalog_masters('amsin_eu_group_by_catalog_masters')
        self.get_all_candidates('amsin_eu_get_all_candidates')
        self.getTestUsersForTest('amsin_eu_getTestUsersForTest')
        # self.interviews('amsin_eu_interviews')
        # self.new_interviews('amsin_eu_interview_new')


Object = AmsinNonEu()
Object.performance_login("amsin_eu")
if Object.login == 'OK':
    for i in range(0, 3):
        Object.api_response_time()
        Object.create_pandas_excel('AMSIN_EU')
        print("Run:: ", i)
