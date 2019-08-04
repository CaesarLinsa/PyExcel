import testtools
import fixtures
import sys
import re
import six
import mock
from PyExcel import shell
from testtools import matchers


class TestBaseShell(testtools.TestCase):

    def setUp(self):
        super(TestBaseShell, self).setUp()
        self.useFixture(fixtures.FakeLogger())

    @mock.patch('sys.stdout', new=six.StringIO())
    def shell(self, argstr):
        try:
            _shell = shell.shellmain()
            _shell.main(argstr.split())
        except SystemExit:
            exc_type, exc_value, exc_traceback = sys.exc_info()
            self.assertEqual(0, exc_value.code)

        return sys.stdout.getvalue()


class TestHelpShell(TestBaseShell):
    RE_OPTIONS = re.DOTALL | re.MULTILINE

    def test_help_command(self):
        required = [
            ".*?^usage: pyexcel",
            ".*?^see pyexcel command for help"
        ]
        for argstr in  ['help', '--help']:
            help_text = self.shell(argstr)
            for r in required:
                self.assertThat(help_text, matchers.MatchesRegex(r,
                                                      self.RE_OPTIONS))
