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

    def __init__(self, run_path, joblist_path, joblist_name):

        self.run_path = run_path
        self.job_path = joblist_path
        self.jbl_name = joblist_name

        self.lc_path = {}  #lc: [pj path, relative path]
        self.lc_list = []  #lc to write

        self.bladed_path   = r'C:\Program Files (x86)\DNV GL\Bladed 4.8\DTBLADED.EXE'
        self.dtbladed_path = r'C:\Program Files (x86)\DNV GL\Bladed 4.8'

        self.get_path()
        self.write_joblist()

    def get_path(self):

        dlc_list = [file for file in os.listdir(self.run_path) if os.path.isdir(os.path.join(self.run_path, file))]

        for dlc in dlc_list:

            dlc_path = os.path.join(self.run_path, dlc)

            lcs_list = [file for file in os.listdir(dlc_path) if os.path.isdir(os.path.join(dlc_path, file))]
           
            for lc in lcs_list:

                lc_path = os.path.join(dlc_path, lc)
                
                file_list = [file.lower() for file in os.listdir(lc_path)]
                pj_file = '.'.join([lc,'$pj'])
                in_file = 'dtbladed.in'

                if pj_file in file_list and in_file in file_list:
                    
                    self.lc_path[lc] = [lc_path, os.path.join(dlc, lc)]
                    self.lc_list.append(lc)

        if not self.lc_list:
            raise Exception('%s\nNo pj and in exist! Please choose a right path!'%self.run_path)

    def write_joblist(self):

        # head
        content = '<?xml version="1.0"?>\n'
        content += '<JobList xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" ' \
                   'xmlns:xsd="http://www.w3.org/2001/XMLSchema>"\n'
        content += '  <JobRuns>\n'

        for index, lc in enumerate(self.lc_list):

            content += '    <JobRun>\n'
            content += '      <Rank>%s</Rank>\n' %(index+1)
            content += '      <Job>\n'
            content += '        <Id ProductName="Bladed" ProductVersion="4.8.0.41">\n'
            content += '          <Text>%s</Text>\n' %lc
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
            content += '        <ResultDir>%s</ResultDir>\n' %self.lc_path[lc][0]
            content += '        <Application>Bladed</Application>\n'
            content += '        <InputFileDir>%s</InputFileDir>\n' %self.lc_path[lc][0]
            content += '        <InputFileDirRelativeToBatch>%s</InputFileDirRelativeToBatch>\n' %self.lc_path[lc][1]
            content += '        <RunInCloud>false</RunInCloud>\n'
            content += '        <BatchRunId />\n'
            content += '        <RunType>RunCalcsToCompletion</RunType>\n'
            content += '        <Version>4.8.0.34</Version>\n'
            content += '        <CloudSessionId />\n'
            content += '        <CompanyId />\n'
            content += '        <Name>%s</Name>\n' %lc
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

        if not os.path.exists(self.job_path):
            os.makedirs(self.job_path)

        joblist_path = os.path.join(self.job_path, self.jbl_name + '.joblist')

        with open(joblist_path, 'w+') as f:
            f.write(content)

if __name__ == '__main__':

    prj_path = r'\\172.20.0.4\fs01\CE\W6250-172-10m\run_failed\0108\DLC72'
    # bld_path = r'C:\Program Files (x86)\DNV GL\Bladed 4.8'
    # bld_path = None
    bat_path = r'\\172.20.0.4\fs01\CE\W6250-172-10m\batch\run_failed'
    # res_path = r'\\172.20.0.4\fs01\CE\W6250-172-10m\run_0108\DLC72'
    jbl_name = '0108'

    # 1:debladed.exe(time series); 2:dtsignal.exe(post); 3:windnd.exe(wind)
    exe_flag = 1

    # bladed version: 4.6 or 4.8
    bld_vers = '4.8'

    Create_Joblist(prj_path, bat_path, jbl_name)