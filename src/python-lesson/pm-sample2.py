# coding:utf-8
 
import paramiko

USER = 'root'
PSWD = 'minh01227'

class InputReader:
    hoge = "aaa"
    def __init__(self, commands_path, hosts_path):
        self.commands_path = commands_path
        self.hosts_path = hosts_path

    def read(self):
        self.commands = self.__readlines(self.commands_path)
        self.hosts = self.__readlines(self.hosts_path)

    def __readlines(self, path):
        with open(path) as f:
            # return map(lambda v: v.strip(), f.readlines())
            return [v.strip() for v in f.readlines()] # リスト内包表記

class CommandExecuter:
    def __init__(self, host, command):
        self.host = host
        self.command = command

    def execute(self):
        ssh = paramiko.SSHClient()
        ssh.set_missing_host_key_policy(paramiko.AutoAddPolicy())
        ssh.connect(self.host, username=USER, password=PSWD)

        stdin, stdout, stderr = ssh.exec_command(self.command)

        errors = stderr.readlines()
        if len(errors) != 0:
            raise Exception(errors)

        lines = [v.strip() for v in stdout.readlines()]
        ssh.close()
        return lines

if __name__ == '__main__':
    reader = InputReader("commands.txt", "hosts.txt")
    reader.read()

    for h in reader.hosts:
        for c in reader.commands:
            executer = CommandExecuter(h, c)
            results = executer.execute()
            print("{0}({1}): {2}".format(h, c, results))