from scripts.performance_testing.output import performance_output


class AmsinNonEu(performance_output.AmsinNonEuOutput):
    def __init__(self):
        self.eu = 'yes'
        self.login_server = 'amsin'
        super(AmsinNonEu, self).__init__()

    def api_response_time(self):
        self.get_tenant_details()
        self.get_all_entity_properties()
        self.group_by_catalog_masters()
        self.get_all_candidates()
        self.getTestUsersForTest()
        # self.interviews()
        # self.new_interviews()


Object = AmsinNonEu()
Object.common_login("amsin_eu")
if Object.login == 'OK':
    loop_run = [1, 2, 3]
    for i in loop_run:
        Object.api_response_time()
        Object.create_pandas_excel('AMSIN_EU')
        print("Run:: ", i)
