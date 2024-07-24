from scripts.performance_testing.output import performance_output


class AmsinNonEu(performance_output.AmsinNonEuOutput):
    def __init__(self):
        self.eu = 'no'
        self.login_server = 'amsin'
        super(AmsinNonEu, self).__init__()

    def api_response_time(self):
        self.get_tenant_details('get_tenant_details')
        self.get_all_entity_properties('get_all_entity_properties')
        self.group_by_catalog_masters('group_by_catalog_masters')
        self.get_all_candidates('get_all_candidates')
        self.getTestUsersForTest('getTestUsersForTest')
        # self.interviews('interviews')
        # self.new_interviews('interview_new')


Object = AmsinNonEu()
Object.performance_login("amsin_non_eu")
if Object.login == 'OK':
    for i in range(0, 3):
        Object.api_response_time()
        Object.create_pandas_excel('AMSIN_NON_EU')
        print("Run:: ", i)
