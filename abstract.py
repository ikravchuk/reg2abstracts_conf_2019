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
# cut last occurance of ")" and split by "),"
registration_data['coauthors'] = registration_data['coauthors'].map(lambda x: x.strip().rstrip(')').split('),'))
# split coauthor's name and affiliation
registration_data['coauthors'] = registration_data['coauthors'].map(lambda x: tuple([i.split('(') for i in x]))

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
    author = '***{} {}.^{}^***'.format(row.last_name_en, row.first_name_en[0], affil_indexes)

    affils = row.affiliation
    affil_output = ''
    for i, a in enumerate(row.affiliation):
        affil_output += '*^{}^{}*<div></div>'.format(i + 1, a)

    coauthors = '**'
    for coauth in row.coauthors:
        if not any(coauth[1] in a for a in affils):
            coauthors += ', ' + '{}^{}^'.format(coauth[0], len(affils) + 1)
            affil_output += '*^{}^{}*<div></div>'.format(len(affils) + 1, coauth[1])
            affils.append(coauth[1])
        else:
            for i, a in enumerate(affils):
                if coauth[1] in a:
                    coauthors += ', ' + '{}^{}^'.format(coauth[0].strip(), i + 1)
    coauthors += "**"
    abstract = '{}'.format(row.abstract)
    
    # markdown compilation
    md += title
    md += '<div></div>'
    md += author + coauthors
    md += '<div></div>'
    md += affil_output
    md += '<div></div>'
    md += abstract
    md += '<div style="page-break-after: always;"></div>'
output = pypandoc.convert_text(md, 'docx', format = 'md', outputfile = 'abstracts.docx')
assert output == ''

#print(registration_data)
