import pandas as pd
import pypandoc

registration_data = pd.read_excel('registration_data.xlsx',
                                  names = ['reg_date',
                                           'email',
                                           'first_name_ua',
                                           'last_name_ua',
                                           'first_name_en',
                                           'last_name_en',
                                           'affiliation',
                                           'coauthors',
                                           'type',
                                           'section',
                                           'additional',
                                           'title_ua',
                                           'title_en',
                                           'abstract'])

# Data Cleaning
registration_data['type'] = registration_data['type'].map(lambda x : x[:x.find(' (')])
registration_data['title_en'] = registration_data['title_en'].map(lambda x : x.upper())

print('Total number: {}'.format(len(registration_data)))
print('Sections :')
print(registration_data['section'].value_counts())
print('Types :')
print(registration_data['type'].value_counts())

# Markdown file for abstract book
md = ''
for row in registration_data.itertuples():
    title = '# {}\n'.format(row.title_en)
    md += title
    md += '<div></div>'
    author = '***{} {}.^1^,***'.format(row.last_name_en, row.first_name_en[0])
    coauthors = '**{}**'.format(row.coauthors)
    md += author + coauthors
    md += '<div></div>'
    abstract = '{}'.format(row.abstract)
    md += abstract
    md += '<div style="page-break-after: always;"></div>'

output = pypandoc.convert_text(md, 'docx', format = 'md', outputfile = 'abstracts.docx')
assert output == ''

#print(registration_data)
