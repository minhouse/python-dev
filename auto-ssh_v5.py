# Modules
import gc
import getpass
import logging
import paramiko
from contextlib import suppress, closing
from paramiko import SSHException

# Variables
CMDLIST = '/home/local-adm/LG_Installation/minho/CR/commands.txt'
SYSLIST = '/home/local-adm/LG_Installation/minho/CR/systems.txt'

# Classes and Functions


class InputReader:
    def __init__(self, commands_path, hosts_path):
        self.commands_path = commands_path
        self.hosts_path = hosts_path

    def read(self):
        self.commands = self.__readlines(self.commands_path)
        self.hosts = self.__readlines(self.hosts_path)

    def __readlines(self, path):
        with open(path) as f:
            return [v.strip() for v in f.readlines()]


class CommandExecuter():
    def __init__(self, host, command, user, pswd):
        self.host = host
        self.command = command
        self.user = user
        self.pswd = pswd

    def execute(self):
        try:
            ssh = paramiko.SSHClient()
            ssh.set_missing_host_key_policy(paramiko.AutoAddPolicy())
            ssh.connect(self.host, username=self.user, password=self.pswd)
            stdin, stdout, stderr = ssh.exec_command(self.command)
            lines = [v.strip() for v in stdout.readlines()]
            ssh.close()
            return lines
        except Exception as err:
            print('[ERROR] %s SSH connection failed' % self.host + '\n')
            raise err
        finally:
            if ssh:
                ssh.close()


def main():
    user = input("Enter User Name: ")
    pswd = getpass.getpass("Enter User Password: ")

    reader = InputReader(CMDLIST, SYSLIST)
    reader.read()

    for h in reader.hosts:
        try:
            for c in reader.commands:
                executer = CommandExecuter(h, c, user, pswd)
                results = executer.execute()
                print("{0} {1}".format(h, c) + '\n')
                if results is not None:
                    for i in results:
                        print(i + '\n')
        except Exception as err:
            logging.exception('%s', err)


# Main Procedure
if __name__ == '__main__':
    main()