#！usr/bin/python3
# -*- coding: utf-8 -*-
# @Time    : 3/9/2020 3:19 PM
# @Author  : CE
# @File    : create_joblist.py
'''
create joblist based on run
'''

import os

class Create_Joblist(object):

    def __init__(self, lc_in_path, lc_res_path, joblist_path, joblist_name):

        self.lc_in_path   = lc_in_path
        self.lc_res_path  = lc_res_path
        self.joblist_name = joblist_name
        self.joblist_path = joblist_path

        self.bladed_path   = r'C:\program files (x86)\dnv gl\bladed 4.8\DTBLADED.exe'
        self.dtbladed_path = r'C:\program files (x86)\dnv gl\bladed 4.8'

        self.lc_list = list(self.lc_in_path.keys())
        self.lc_list.sort()

        self.write_joblist()

    def write_joblist(self):

        # head
        content = '<?xml version="1.0"?>\n'
        content += '<JobList xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" ' \
                   'xmlns:xsd="http://www.w3.org/2001/XMLSchema">\n'
        content += '  <JobRuns>\n'

        for index, lc in enumerate(self.lc_list):

            relative_path = os.sep.join(self.lc_in_path[lc].replace('/', '\\').split(os.sep)[-2:])

            content += '    <JobRun>\n'
            content += '      <Rank>%s</Rank>\n' %(index+1)
            content += '      <Job>\n'
            content += '        <Id ProductName="Bladed" ProductVersion="4.8.0.41">\n'
            content += '          <Text>%s</Text>\n' %lc.split(os.sep)[-1]
            content += '        </Id>\n'
            content += '        <CalculationParameters>\n'
            content += '          <ProjectFilePath />\n'
            content += '          <ExecutableName>%s</ExecutableName>\n' %self.bladed_path
            content += '          <ExecutableVersion />\n'
            content += '          <ExecutableDirectory>%s</ExecutableDirectory>\n' %self.dtbladed_path
            content += '          <InputFiles>\n'
            content += '            <InputFile>\n'
            content += '              <FileName>-</FileName>\n'
            content += '              <Content />\n'
            content += '            </InputFile>\n'
            content += '          </InputFiles>\n'
            content += '          <CommandLineArgs>-u </CommandLineArgs>\n'
            content += '          <WorkingDirectory />\n'
            content += '          <UsesMappedExecutable>false</UsesMappedExecutable>\n'
            content += '        </CalculationParameters>\n'
            content += '        <ResultDir>%s</ResultDir>\n' %self.lc_res_path[lc].replace('/','\\')
            content += '        <Application>Bladed</Application>\n'
            content += '        <InputFileDir>%s</InputFileDir>\n' %self.lc_in_path[lc].replace('/','\\')
            content += '        <InputFileDirRelativeToBatch>%s</InputFileDirRelativeToBatch>\n' %relative_path.replace('/','\\')
            content += '        <RunInCloud>false</RunInCloud>\n'
            content += '        <BatchRunId />\n'
            content += '        <RunType>RunCalcsToCompletion</RunType>\n'
            content += '        <Version>4.8.0.34</Version>\n'
            content += '        <CloudSessionId />\n'
            content += '        <CompanyId />\n'
            content += '        <Name>%s</Name>\n' %lc.split(os.sep)[-1]
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
        content += '    <Text>76c32621-904e-4254-a9ca-43341c616186</Text>\n'
        content += '  </Id>\n'
        content += '</JobList>'

        if not os.path.isdir(self.joblist_path):
            os.makedirs(self.joblist_path)

        joblist_name = self.joblist_name + '.joblist'
        joblist_path = os.path.join(self.joblist_path, joblist_name)

        with open(joblist_path, 'w+') as f:
            f.write(content)

if __name__ == '__main__':

    prj_path = r'\\172.20.4.132\fs02\CE\V3\loop06\performance\ct'

    lc_path = {}
    for file in os.listdir(prj_path):
        if file!='model':
            print(file)
            lc_path[file] = os.path.join(prj_path, str(file))
    print(lc_path)
    joblist = 'ct'
    jl_path = r'\\172.20.4.132\fs02\CE\V3\loop06\performance\ct'

    Create_Joblist(lc_path, lc_path, jl_path, joblist)