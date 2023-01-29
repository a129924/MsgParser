# eml_writer
from email.mime.text import MIMEText
from email.generator import Generator
from email.mime.application import MIMEApplication
from email.mime.multipart import MIMEMultipart

# msg_reader
import os
from extract_msg import message


class MsgFileParser:
    def __init__(self, msg_file: str, to_path: str = r".\\"):
        self._msg_file = msg_file
        self.to_path = to_path

    @property
    def msg_file(self):
        return self._msg_file

    @msg_file.setter
    def msg_file(self, value):
        self._msg_file = value


class MsgFileReader(MsgFileParser):
    def __init__(self, msg_file: str, to_path: str = r".\\",
                 need_extract_attachment: bool = False,
                 need_extract_mail_body: bool = False):

        super(MsgFileReader, self).__init__(msg_file, to_path)
        self.need_extract_attachment = need_extract_attachment
        self.need_extract_mail_body = need_extract_mail_body

    @property
    def msg(self): return message.Message(self._msg_file)

    @property
    def attachments(self): return self.msg.attachments

    @property
    def body(self) -> str: return str(self.msg.body)

    @property
    def subject(self): return self.msg.subject

    @property
    def sender_email(self): return self.msg.sender

    @property
    def CC(self) -> list: return self.split_string(self.msg.cc)

    @property
    def BCC(self) -> list: return self.split_string(self.msg.bcc)

    @property
    def to(self) -> list: return self.split_string(self.msg.to)

    @staticmethod
    def split_string(string) -> list:
        if isinstance(string, str) and ";" in string:
            return string.split(";")

        return [string]

    def extract_attachments(self, to_path: str = "") -> None:
        assert self.need_extract_attachment == True and self.attachments
        to_path = os.path.join(os.getcwd(), to_path) if to_path != "" else os.path.join(
            os.getcwd(), self.to_path)
        if os.path.exists(to_path) is False:
            os.mkdir(to_path)

        for attachment in self.attachments:
            print(attachment.getFilename()) # 附件檔案名稱
            print(attachment.data) # 附件檔案內容
        
    def save_email_body_to_txt(self, filename: str, output_path: str = ".\\"):
        assert self.body

        if os.path.exists(output_path) is False:
            os.mkdir(output_path)

        fullpath = os.path.join(os.getcwd(), output_path, filename)

        with open(fullpath, 'w+',  encoding='utf-8') as writer:
            writer.write(self.body)

    class EmlWriter():
        def __init__(self, filename: str, subject: str, From: str,
                     To: str, content: str, attachments: list = [],
                     src_path: str = ".\\", to_path: str = ".\\", content_type="string"):
            assert content_type in ["string, file"]

            self._filename = filename if os.path.splitext(
                filename)[1][1:] == ".eml" else filename.join(".eml")
            self._subject = subject
            self._From = From
            self._To = To
            self._content = content
            self._attachments = attachments
            self._src_path = src_path
            self._to_path = to_path
            self._content_type = content_type

        @property
        def content(self) -> str:
            if self._content_type == "file":
                assert os.path.isfile(self._content)
                with open(self._content, "r") as f:
                    text = f.read()

                return text
            else:
                return self._content

        @property
        def attachments_all_exists(self) -> bool:
            if self._attachments == []:
                return False

            return all(list(map(os.path.isfile, self._attachments)))

        @property
        def msg(self) -> MIMEMultipart:
            return MIMEMultipart()

        @property
        def output_full_path(self):
            return os.path.join(self._to_path, self._filename)

        def add_attachment_to_email(self,):
            if self.attachments_all_exists:
                for attachment in self._attachments:
                    out_file = MIMEApplication(
                        open(os.path.join(self._src_path, attachment), "rb").read())
                    out_file.add_header('Content-Disposition',
                                        'attachment', filename=attachment)
                    self.msg.attach(out_file)

        def output_eml(self,):
            self.msg["Subject"] = self._subject
            self.msg["From"] = self._From
            self.msg["To"] = self._To

            self.msg.attach(MIMEText(self._content, 'plain', 'utf-8'))
            self.add_attachment_to_email()

            with open(self._filename, 'w') as out:

                gen = Generator(out)
                gen.flatten(self.msg)


if __name__ == '__main__':
    msg_reader = MsgFileReader(
        'RESimplemessage.msg', need_extract_attachment=True)
    print(msg_reader.attachments)
    msg_reader.extract_attachments(to_path = "DATA")
