import re
import pandas as pd
import docx


class ResumeParser(object):
    experience = ('experience', 'background', 'employment',)
    education = ('education', 'academic',)
    skills = ('skills', 'qualifications', 'knowledge', 'competencies',)
    summary = ('summery', 'objective', 'career',)
    accomplishments = ('projects', 'publications',)
    extra_activity = ('intrests', 'volunteer', 'honors')

    def __init__(self, file_name):
        self.partitions = {'experience': {}, 'summary': {}, 'skills': {}, 'education': {}, 'contact_info': {},
                           'accomplishments': {},
                           'extra_activity': {}, }
        self.email = None
        self.contact_frame = None
        self.work_frame = None
        self.education_frame = None
        self.skills_frame = None
        self.file_name = file_name
        self.text = None
        self.read_file()
        self.section_index = list()
        self.find_partitions()
        self.extract_partition()
        self.contact_info()
        self.create_work_frame()
        self.create_education_frame()
        self.create_skill()

    def read_file(self):
        if self.file_name.endswith('.txt'):
            with open(self.file_name, 'r') as file:
                self.text = file.readlines()
        elif file_name.endswith('.docx'):
            document = docx.Document(file_name)
            self.text = []
            for para in document.paragraphs:
                self.text.append(para.text)

    def find_partitions(self):
        """This function recive the resume text as input and use the partition dictionary for finding
        the key word for different section and return a list containing the starting point of each section"""
        for index, line in enumerate(self.text):
            if len(line) == 0:
                continue
            if line[0].islower():
                continue

            line = line.lower().strip()
            if [word for word in ResumeParser.summary if word in line]:
                self.partitions['summary'][line] = index
                self.section_index.append(index)
            elif [word for word in ResumeParser.experience if word in line]:
                self.partitions['experience'][line] = index
                self.section_index.append(index)
            elif [word for word in ResumeParser.skills if word in line]:
                self.partitions['skills'][line] = index
                self.section_index.append(index)
            elif [word for word in ResumeParser.education if word in line]:
                self.partitions['education'][line] = index
                self.section_index.append(index)
            elif [word for word in ResumeParser.accomplishments if word in line]:
                self.partitions['accomplishments'][line] = index
                self.section_index.append(index)
            elif [word for word in ResumeParser.extra_activity if word in line]:
                self.partitions['extra_activity'][line] = index
                self.section_index.append(index)

    def extract_partition(self):
        """This function recive the resume text and extarct and store each section in partitions variable using
        the list of indexes created by find_partions function """

        self.partitions['contact_info'] = self.text[:self.section_index[0]]

        for section, value in self.partitions.items():
            if section == 'contact_info':
                continue

            for sub_section, str_index in value.items():
                end_index = len(self.text)
                if (self.section_index.index(str_index) + 1) != len(self.section_index):
                    end_index = self.section_index[self.section_index.index(str_index) + 1]
                self.partitions[section][sub_section] = self.text[str_index:end_index]

    def contact_info(self):
        toeknz = self.partitions['contact_info'][0].split()
        name, family = toeknz[0], toeknz[1]
        self.partitions['contact_info'] = " ".join(self.partitions['contact_info'])
        try:
            phone = re.search(r'(\+\d{1,2}[\s.-])?\(?\d{3}\)?[\s.-]\d{3}[\s.-]\d{4}',
                              self.partitions['contact_info']).group()
        except:
            phone = ''

        email = re.findall(r"([^@|\s]+@[^@]+\.[^@|\s]+)", self.partitions['contact_info'])
        if email:
            try:
                email = email[0].split()[0].strip(';')
            except IndexError:
                email = ''

        self.email = str(email)
        self.contact_frame = pd.DataFrame({'Name': name + ' ' + family, 'Email': email, 'Phone Number': phone},
                                          index=[name])

    def create_skill(self):
        skills = []
        for key, value in self.partitions['skills'].items():
            for item in value:
                item = item.replace('\n', '').strip()
                if key in item.lower():
                    continue
                skills.extend(re.split(':|,', item)[1:])

        self.skills_frame = pd.DataFrame(skills, columns=['skills'], index=[self.email for item in skills])

    @staticmethod
    def check_work(line):
        check = []
        for word in line:
            if word.isdigit():
                continue
            check.append(word.isupper() or word.istitle())
        return all(check)

    def create_work_frame(self):
        work = []
        for key, value in self.partitions['experience'].items():
            for line in value:
                if key in line.strip().lower():
                    continue
                tokens = [word for word in re.split(',|\s|-', line.strip())[1:] if word != '']
                if ResumeParser.check_work(tokens):
                    work.append(list(tokens))
        company = work[:-1:2]
        company = [' '.join(comp) for comp in company]
        position = work[1::2]
        position = [' '.join(pos) for pos in position]
        self.work_frame = pd.DataFrame({'Company': company, 'Position': position},
                                       index=[self.email for item in company])

    def create_education_frame(self):
        university = []
        degree = []
        for key, value in self.partitions['education'].items():
            for line in value:

                if key in line.strip().lower():
                    continue
                tokens = [word.lower() for word in re.split(',|\s|-', line.strip())[:] if word != '']
                if tokens and tokens[0] == 'â€¢':
                    tokens.remove('â€¢')
                if 'university' in tokens:
                    university.append(tokens)
                    continue
                if 'master' in tokens or 'bachelor' in tokens or 'doctorate' in tokens:
                    degree.append(tokens)

        university = [' '.join(item) for item in university]
        degree = [' '.join(item).replace('{', '-') for item in degree]
        self.education_frame = pd.DataFrame({'University': university, 'Degree': degree},
                                            index=[self.email for item in university])

    def get_dataframes(self):
        return self.contact_frame, self.work_frame, self.education_frame, self.skills_frame


file_name = 'resume.docx'
contact_frame, work_frame, education_frame, skills_frame = ResumeParser(file_name).get_dataframes()
print('='*100)
print(contact_frame)
print('='*100)
print(work_frame)
print('='*100)
print(education_frame)
print('='*100)
print(skills_frame)
