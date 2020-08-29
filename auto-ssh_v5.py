# Modules
import gc
import getpass
import paramiko
from contextlib import suppress, closing
from paramiko import SSHException

# Variables
# USER = 'root'
# PSWD = 'P@ssw0rd1234'

# Classes and Functions


class InputLoginCredential:
    '''スクリプト実行時にログイン先のユーザー名とパスワードを入力させるためのクラス'''
    def inputcrendential(self):
        self.user = input("Enter User Name: ")
        self.pswd = getpass.getpass("Enter User Password: ")


class InputReader:
    '''リモート先で実行するコマンドのリストとログイン先のIPリストのファイルを読み込むためのクラス'''
    def __init__(self, commands_path, hosts_path):
        self.commands_path = commands_path
        self.hosts_path = hosts_path

    def read(self):
        self.commands = self.__readlines(self.commands_path)
        self.hosts = self.__readlines(self.hosts_path)

    def __readlines(self, path):
        with open(path) as f:
            return [v.strip() for v in f.readlines()]


class CommandExecuter(InputLoginCredential):
    '''sshのログイン処理とコマンドを実行するためのクラス
    InputLoginCredentialクラスで入力されたuserとpswdを利用する'''
    def __init__(self, host, command, user, pswd):
        self.host = host
        self.command = command
        self.user = user
        self.pswd = pswd

    def execute(self):
        with suppress(Exception):
            with closing(paramiko.SSHClient()) as ssh:
                ssh.set_missing_host_key_policy(paramiko.AutoAddPolicy())
                '''InputLoginCredentialクラスを定義する前は上に定義した
                USERとPSWDの変数を以下のusernameとpasswordに渡していた
                ssh.connect(self.host, username=USER, password=PSWD)'''
                ssh.connect(self.host, username=self.user, password=self.pswd)
                stdin, stdout, stderr = ssh.exec_command(self.command)
                errors = stderr.readlines()
                lines = [v.strip() for v in stdout.readlines()]
                return lines
        print('## %s SSH connection failed ##' % self.host + '\n')

    def __del__(self):
        self.host = None
        self.command = None
        self.commands_path = None
        self.hosts_path = None
        self.exec_command = None
        del self


def main():
    '''インスタンスの生成'''
    reader = InputReader("/root/CR/commands.txt", "/root/CR/systems.txt")
    reader.read()

    for h in reader.hosts:
        for c in reader.commands:
            executer = CommandExecuter(h, c)
            results = executer.execute()
            print("IP: {0} :({1}):".format(h, c) + '\n')
            if results is not None:
                for i in results:
                    print(i + '\n')
        del executer
        del results
        gc.collect()


# Main Procedure
if __name__ == '__main__':
    main()