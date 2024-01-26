import email
from email.message import Message
import sys
import os
import quopri
from pathlib import Path
from ansi_colors import color


def info(msg):
    print(color(msg, fg='black', bg='blue'))


def success(msg):
    print(color(msg, fg='green'))


def error(msg):
    print(color(msg, fg='red'))


class FindEmailParts:
    def __init__(self, file_path: Path):
        self.plain_counter: int = 0
        self.parent = file_path.parent
        self.msg = file_path

    @property
    def msg(self) -> Message:
        return self._msg

    @msg.setter
    def msg(self, val):
        with open(val, "r") as fr:
            msg = email.message_from_file(fr)
            self._msg = msg

    def _extract_part(self, part: Message):
        """ This will extract the every part depending on type """

        sub_msg: Message
        for msg_id, sub_msg in enumerate(part.get_payload()):
            info(f"Message: {sub_msg.get_content_type()}")
            if sub_msg.is_multipart():
                self._extract_part(sub_msg)
            elif sub_msg.get_content_disposition() == 'attachment':
                self._extract_attachment(sub_msg)
            elif sub_msg.get_content_type() == 'text/plain':
                self._extract_text(sub_msg, extension='txt')
            elif sub_msg.get_content_type() == 'text/html':
                self._extract_text(sub_msg, extension='html')
            elif sub_msg.get_content_type().startswith('image'):
                self._extract_image(sub_msg)
            else:
                error(f'{msg_id}: Cannot be parsed {sub_msg.get_content_type()}')

    def _extract_image(self, image_message: Message):
        data = image_message.get_payload(decode=True)
        fp = self.parent / image_message.get_filename()

        with open(fp, 'wb') as f:
            f.write(data)
        success(f'Saved image: {fp}')

    def _extract_text(self, plain_message: Message, extension: str):
        if "Content-Transfer-Encoding" in plain_message.keys():
            payload_msg = quopri.decodestring(plain_message.get_payload()).decode(plain_message.get_content_charset())
        else:
            assert False

        if payload_msg:
            self.plain_counter += 1
            fp = self.parent / f'plain{self.plain_counter}.{extension}'
            with open(fp, 'w') as f:
                f.write(payload_msg)
            success(f'Saved document: {fp}')

    def _extract_attachment(self, attachment: Message):
        fp = self.parent / attachment.get_filename()
        with open(fp, "wb") as fw:
            file = attachment.get_payload(decode=True)
            fw.write(file)
            success(f'Saved attachment: {fp}')

    def retrieve_parts(self):
        if self.msg.is_multipart():
            info(f"Message: {self.msg.get_content_type()}")
            self._extract_part(self.msg)


if __name__ == "__main__":
    if len(sys.argv) != 2:
        exit(f"Usage: {os.path.basename(__file__)} <MIME file>")

    p = Path(sys.argv[1])

    if not p.exists() or not p.is_file():
        exit(f"Defined input does not seem to exist or is not a file: {p}")

    fem = FindEmailParts(p)
    fem.retrieve_parts()
