#!/usr/bin/env python
TEST_STATUSES = {'pass' : 'Pass',
                 'fail' : 'Fail',
                 'warning' : 'Warning',
                 'skip' : 'Skip'}

TEST_STATUS_PRIORITIES = {
                    '': 0,
                    None: 0,
                    'Skip': 1,
                    'Pass': 2,
                    'Warning': 3,
                    'Fail': 4 }
