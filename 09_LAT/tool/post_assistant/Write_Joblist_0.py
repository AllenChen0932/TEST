#！usr/bin/python3
# -*- coding: utf-8 -*-
# @Time    : 3/9/2020 3:19 PM
# @Author  : CE
# @File    : Write_Joblist.py

import os

class Create_Joblist(object):

    def __init__(self, post_path, joblist_name, joblist_path):

        self.post_path = post_path.replace('/', '\\')
        self.jobl_name = joblist_name
        self.jobl_path = joblist_path
        
        self.post_list = []
        self.upj_path  = {}

        self.bladed_path = r'c:\program files (x86)\dnv gl\bladed 4.8'
    
        self.write_post_joblist()

    def write_post_joblist(self):

        for root, dirs, files in os.listdir(self.post_path):

            for file in files:

                if '$PJ' in file and 'dtsignal.in' in files:

                    pj_name = file.split('.$')[0]
                    pj_path = os.path.join(root, file)

                    self.post_list.append(pj_name)

                    self.upj_path[pj_name] = pj_path.replace(' ', '_')
        # print(self.post_list)

        # head
        content = '<?xml version="1.0"?>\n'
        content += '<JobList xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" ' \
                   'xmlns:xsd="http://www.w3.org/2001/XMLSchema">\n'
        content += '  <JobRuns>\n'

        for index, pj in enumerate(self.post_list):

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

        if not os.path.exists(self.jobl_path):

            os.makedirs(self.jobl_path)

        joblist_name = '.'.join((self.jobl_name,'joblist'))
        joblist_path = os.path.join(self.jobl_path, joblist_name)

        with open(joblist_path, 'w+') as f:

            f.write(content)

if __name__ == '__main__':

    post_path = r'\\172.20.0.4\fs03\CE\post_test\test'

    Create_Joblist(post_path=post_path, joblist_name='post')