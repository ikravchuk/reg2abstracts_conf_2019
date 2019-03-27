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
registration_data['affiliation'] = registration_data['affiliation'].map(lambda x: x.split('|'))
print(registration_data['affiliation'])
#registration_data['coauthors'] = registration_data['coauthors'].map(lambda x: x.split('),'))

print('Total number: {}'.format(len(registration_data)))
print('Sections :')
print(registration_data['section'].value_counts())
print('Types :')
print(registration_data['type'].value_counts())

# Markdown file for abstract book
md = ''
for row in registration_data.itertuples():
    # formation of text blocks in markdown
    title = '# {}\n'.format(row.title_en)
    affil_indexes = ','.join(str(n) for n in range(1, len(row.affiliation) + 1)) # affiliation indicators for author
    author = '***{} {}.^{}^,*** '.format(row.last_name_en, row.first_name_en[0], affil_indexes)


    #coauthors = '**{}**'.format(row.coauthors)
    affiliation = ''
    for i, a in enumerate(row.affiliation):
        affiliation += '*^{}^{}*<div></div>'.format(i + 1, a)

    abstract = '{}'.format(row.abstract)
    
    # markdown compilation
    md += title
    md += '<div></div>'
    md += author #+ coauthors
    md += '<div></div>'
    md += affiliation
    md += '<div></div>'
    md += abstract
    md += '<div style="page-break-after: always;"></div>'
print(md)
output = pypandoc.convert_text(md, 'docx', format = 'md', outputfile = 'abstracts.docx')
assert output == ''

#print(registration_data)
