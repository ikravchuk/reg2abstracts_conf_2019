import pandas as pd

registration_data = pd.read_excel('registration_data.xlsx',
                                  names = ['reg_date',
                                           'email',
                                           'first_name_ua',
                                           'last_name_ua',
                                           'first_name_en',
                                           'last_name_en',
                                           'affiliation',
                                           'co-authors',
                                           'type',
                                           'section',
                                           'additional',
                                           'title_ua',
                                           'title_en',
                                           'abstract'])

registration_data['type'] = registration_data['type'].map(lambda x : x[:x.find(' (')])
print('Total number: {}'.format(len(registration_data)))
print('Sections :')
print(registration_data['section'].value_counts())
print('Types :')
print(registration_data['type'].value_counts())
print(registration_data)
