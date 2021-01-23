#！usr/bin/python3
# -*- coding: utf-8 -*-
# @Time    : 3/9/2020 3:19 PM
# @Author  : CE
# @File    : create_joblist.py

import os

class Create_Joblist(object):

    def __init__(self,
                 post_path,
                 ult_identifier=None,
                 fat_identifier=None,
                 ultimate_option=None,
                 fatigue_option=None):

        self.post_path = post_path.replace('/', '\\')
        self.ult_ident = ult_identifier
        self.fat_ident = fat_identifier
        self.ultimate  = ultimate_option
        self.fatigue   = fatigue_option

        self.ult_path = os.path.join(self.post_path, ('ultimate_'+self.ult_ident)) \
            if self.ult_ident else os.path.join(self.post_path, 'ultimate')
        print(self.ult_path)
        self.upj_list = []   # ultimate project name list
        self.upj_path = {}   # ultimate project: path

        self.fat_path = os.path.join(self.post_path, ('fatigue_'+self.fat_ident))\
            if self.fat_ident else os.path.join(self.post_path, 'fatigue')
        print(self.fat_path)
        self.fpj_list = []
        self.fpj_path = {}

        # joblist path
        self.jblist_path = os.path.join(self.post_path, 'Job Lists')

        self.bladed_path = r'c:\program files (x86)\dnv gl\bladed 4.8'

        # self.get_pj()
        if self.ultimate:
            print('begin to write ultimate joblist')
            self.write_ultimate_joblist()
            self.ult_flag = True
        else:
            self.ult_flag = False

        if self.fatigue:
            print('begin to write fatigue joblist')
            self.write_fatigue_joblist()
            self.fat_flag = True
        else:
            self.fat_flag = False
    '''
    def get_pj(self):

        file_list = os.listdir(self.ult_path)
        print(file_list)

        for file in file_list:

            if file.lower() == 'ultimate':

                file_path = os.path.join(self.post_path, file)
                for root, dirs, files in os.walk(file_path):

                    for file in files:

                        if '$PJ' in file and 'dtsignal.in' in files:

                            pj_name = file.split('.$')[0]

                            self.upj_list.append(pj_name)

                            self.upj_path[pj_name] = root.replace(' ', '_')

        file_list = os.listdir(self.fat_path)
        for file in file_list:

            if file.lower() == 'fatigue':

                file_path = os.path.join(self.post_path, file)
                for root, dirs, files in os.walk(file_path):

                    for file in files:

                        if '$PJ' in file and 'dtsignal.in' in files:

                            pj_name = file.split('.$')[0]

                            self.fpj_list.append(pj_name)

                            self.fpj_path[pj_name] = root.replace(' ', '_')

        # self.upj_list.sort()
        # self.fpj_list.sort()
        print(self.upj_list, self.fpj_list)
    '''
    def write_ultimate_joblist(self):

        # get ultimate project list
        file_list = os.listdir(self.ult_path)
        print(file_list)
        for file in file_list:

            pj_path = os.path.join(self.ult_path, file)
            files = os.listdir(pj_path)

            for file in files:

                if '$PJ' in file and 'dtsignal.in' in files:

                    pj_name = file.split('.$')[0]

                    self.upj_list.append(pj_name)

                    self.upj_path[pj_name] = pj_path.replace(' ', '_')
        # print(self.upj_list)

        # head
        content = '<?xml version="1.0"?>\n'
        content += '<JobList xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" ' \
                   'xmlns:xsd="http://www.w3.org/2001/XMLSchema">\n'
        content += '  <JobRuns>\n'

        for index, pj in enumerate(self.upj_list):

            relative_path = os.sep.join(self.upj_path[pj].split(os.sep)[-2:])
            pj_file = self.upj_path[pj].split(os.sep)[-1]

            content += '    <JobRun>\n'
            content += '      <Rank>%s</Rank>\n' %(index+1)
            content += '      <Job>\n'
            content += '        <Id ProductName="Bladed" ProductVersion="4.8.0.41">\n'
            content += '          <Text>%s</Text>\n' %pj_file
            content += '        </Id>\n'
            content += '        <CalculationParameters>\n'
            content += '          <ProjectFilePath />\n'
            content += '          <ExecutableName>%s</ExecutableName>\n' %os.sep.join([self.bladed_path, 'DTSIGNAL.exe'])
            content += '          <ExecutableVersion />\n'
            content += '          <ExecutableDirectory>%s</ExecutableDirectory>\n' %self.bladed_path
            content += '          <InputFiles>\n'
            content += '            <InputFile>\n'
            content += '              <FileName>-</FileName>\n'
            content += '              <Content />\n'
            content += '            </InputFile>\n'
            content += '          </InputFiles>\n'
            content += '          <CommandLineArgs />\n'
            content += '          <WorkingDirectory />\n'
            content += '          <UsesMappedExecutable>false</UsesMappedExecutable>\n'
            content += '        </CalculationParameters>\n'
            content += '        <ResultDir>%s</ResultDir>\n' %self.upj_path[pj]
            content += '        <Application>Bladed</Application>\n'
            content += '        <InputFileDir>%s</InputFileDir>\n' %self.upj_path[pj]
            content += '        <InputFileDirRelativeToBatch>%s</InputFileDirRelativeToBatch>\n' %relative_path
            content += '        <RunInCloud>false</RunInCloud>\n'
            content += '        <BatchRunId />\n'
            content += '        <RunType>RunCalcsToCompletion</RunType>\n'
            content += '        <Version>4.8.0.34</Version>\n'
            content += '        <CloudSessionId />\n'
            content += '        <CompanyId />\n'
            content += '        <Name>%s</Name>\n' %pj.split('.$')[0]
            content += '      </Job>\n'
            content += '      <JobStatus>\n'
            content += '        <Status>Queued</Status>\n'
            content += '        <Messages />\n'
            content += '        <HasTruncatedMessages>false</HasTruncatedMessages>\n'
            content += '        <MessagesFileHref />\n'
            content += '        <AdditionalStatusInformation />\n'
            content += '      </JobStatus>\n'
            content += '      <Timing>\n'
            content += '        <StartTime>0001-01-01T00:00:00</StartTime>\n'
            content += '        <EndTime>0001-01-01T00:00:00</EndTime>\n'
            content += '      </Timing>\n'
            content += '      <Progress>0</Progress>\n'
            content += '      <Enabled>true</Enabled>\n'
            content += '      <RunnerName />\n'
            content += '      <ShouldRetry>false</ShouldRetry>\n'
            content += '      <Retried>false</Retried>\n'
            content += '      <RetrialMessage />\n'
            content += '      <RunByVersion />\n'
            content += '      <AbortMode>AbortAllNoDownload</AbortMode>\n'
            content += '    </JobRun>\n'
        content += '  </JobRuns>\n'

        # end
        content += '  <Id ProductName="Batch" ProductVersion="1.4.0.57">\n'
        content += '    <Text>78397c35-36fa-46ae-9d0d-bfd81740752d</Text>\n'
        content += '  </Id>\n'
        content += '</JobList>\n'

        if not os.path.exists(self.jblist_path):

            os.makedirs(self.jblist_path)

        ult_name = 'ultimate%s.joblist' % (('_' + self.ult_ident) if self.ult_ident else '')
        joblist_path = os.path.join(self.jblist_path, ult_name)

        with open(joblist_path, 'w+') as f:

            f.write(content)

    def write_fatigue_joblist(self):

        # get fatigue project
        file_list = os.listdir(self.fat_path)
        print(file_list)
        for file in file_list:

            pj_path = os.path.join(self.fat_path, file)
            files = os.listdir(pj_path)

            for file in files:

                if '$PJ' in file and 'dtsignal.in' in files:

                    pj_name = file.split('.$')[0]

                    self.fpj_list.append(pj_name)

                    self.fpj_path[pj_name] = pj_path.replace(' ', '_')

        # head
        content = '<?xml version="1.0"?>\n'
        content += '<JobList xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" ' \
                   'xmlns:xsd="http://www.w3.org/2001/XMLSchema">\n'
        content += '  <JobRuns>\n'

        for index, pj in enumerate(self.fpj_list):

            relative_path = os.sep.join(self.fpj_path[pj].split(os.sep)[-2:])
            pj_file = self.fpj_path[pj].split(os.sep)[-1]

            content += '    <JobRun>\n'
            content += '      <Rank>%s</Rank>\n' %(index+1)
            content += '      <Job>\n'
            content += '        <Id ProductName="Bladed" ProductVersion="4.8.0.41">\n'
            content += '          <Text>%s</Text>\n' %pj_file
            content += '        </Id>\n'
            content += '        <CalculationParameters>\n'
            content += '          <ProjectFilePath />\n'
            content += '          <ExecutableName>%s</ExecutableName>\n' %os.sep.join([self.bladed_path, 'DTSIGNAL.exe'])
            content += '          <ExecutableVersion />\n'
            content += '          <ExecutableDirectory>%s</ExecutableDirectory>\n' %self.bladed_path
            content += '          <InputFiles>\n'
            content += '            <InputFile>\n'
            content += '              <Content />\n'
            content += '              <FileName>-</FileName>\n'
            content += '            </InputFile>\n'
            content += '          </InputFiles>\n'
            content += '          <CommandLineArgs />\n'
            content += '          <WorkingDirectory />\n'
            content += '          <UsesMappedExecutable>false</UsesMappedExecutable>\n'
            content += '        </CalculationParameters>\n'
            content += '        <ResultDir>%s</ResultDir>\n' %self.fpj_path[pj]
            content += '        <Application>Bladed</Application>\n'
            content += '        <InputFileDir>%s</InputFileDir>\n' %self.fpj_path[pj]
            content += '        <InputFileDirRelativeToBatch>%s</InputFileDirRelativeToBatch>\n' %relative_path
            content += '        <RunInCloud>false</RunInCloud>\n'
            content += '        <BatchRunId />\n'
            content += '        <RunType>RunCalcsToCompletion</RunType>\n'
            content += '        <Version>4.8.0.34</Version>\n'
            content += '        <CloudSessionId />\n'
            content += '        <CompanyId />\n'
            content += '        <Name>%s</Name>\n' %pj.split('.$')[0]
            content += '      </Job>\n'
            content += '      <JobStatus>\n'
            content += '        <Status>Queued</Status>\n'
            content += '        <Messages />\n'
            content += '        <HasTruncatedMessages>false</HasTruncatedMessages>\n'
            content += '        <MessagesFileHref />\n'
            content += '        <AdditionalStatusInformation />\n'
            content += '      </JobStatus>\n'
            content += '      <Timing>\n'
            content += '        <StartTime>0001-01-01T00:00:00</StartTime>\n'
            content += '        <EndTime>0001-01-01T00:00:00</EndTime>\n'
            content += '      </Timing>\n'
            content += '      <Progress>0</Progress>\n'
            content += '      <Enabled>true</Enabled>\n'
            content += '      <RunnerName />\n'
            content += '      <ShouldRetry>false</ShouldRetry>\n'
            content += '      <Retried>false</Retried>\n'
            content += '      <RetrialMessage />\n'
            content += '      <RunByVersion />\n'
            content += '      <AbortMode>AbortAllNoDownload</AbortMode>\n'
            content += '    </JobRun>\n'
        content += '  </JobRuns>\n'

        # end
        content += '  <Id ProductName="Batch" ProductVersion="1.4.0.57">\n'
        content += '    <Text>78397c35-36fa-46ae-9d0d-bfd81740752d</Text>\n'
        content += '  </Id>\n'
        content += '</JobList>\n'

        if not os.path.exists(self.jblist_path):

            os.makedirs(self.jblist_path)

        fat_name = 'fatigue%s.joblist' %(('_'+self.fat_ident) if self.fat_ident else '')
        joblist_path = os.path.join(self.jblist_path, fat_name)

        with open(joblist_path, 'w+') as f:

            f.write(content)

if __name__ == '__main__':

    post_path = r'\\172.20.0.4\fs03\CE\post_test\test'

    Create_Joblist(post_path=post_path,
                   ult_identifier='1',
                   fat_identifier='1',
                   ultimate_option=True,
                   fatigue_option=True)