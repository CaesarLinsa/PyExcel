import sys
import argparse
from util import args
from exc import CommandError


class shellmain(object):

    def get_base_parse(self):
        parser = argparse.ArgumentParser(
            prog='pyexcel',
            description='',
            epilog='see pyexcel command for help',
            add_help=False
        )

        parser.add_argument('-v','--version',
                            action='version',
                            version='1.0'
                            )
        parser.add_argument('-h','--help',
                            action='store_true',
                            help=argparse.SUPPRESS
                            )

        return parser

    def import_modules(self, path):
        __import__(path)
        modules = sys.modules[path]
        return modules

    def _find_action(self, subparsers, sub_modules):
        for fn_name in (func for func in dir(sub_modules) if func.startswith('do_')):
            command = fn_name[3:].replace('_','-')
            callback = getattr(sub_modules,fn_name)
            desc=callback.__doc__ or ''
            help=desc.strip()
            arguments = getattr(callback,'arguments',[])
            subparser = subparsers.add_parser(
                                             command,
                                             help=help,
                                             description=desc,
                                             add_help=False
                                            )
            subparser.add_argument('-h', '--help', action='help',
                                               help=argparse.SUPPRESS)
            self.subcommands[command] = subparser
            for (args,kwargs) in arguments:
                subparser.add_argument(*args, **kwargs)
            subparser.set_defaults(func=callback)

    def get_subcommand_parser(self):
        # keep all the subcommand parsers
        self.subcommands = {}
        parser = self.get_base_parse()
        subparser= parser.add_subparsers(metavar='<subcommand>')
        sub_modules = self.import_modules('PyExcel.PyExcel')
        self._find_action(subparser, sub_modules)
        # add this module's do_help
        self._find_action(subparser, self)
        return parser
    

    @args('command', metavar='<subcommand>', nargs='?',
               help='Display help for <subcommand>')
    def do_help(self, args):
        """Display help about this program or one of its subcommands."""
        if getattr(args, 'command', None):
            if args.command in self.subcommands:
                self.subcommands[args.command].print_help()
            else:
                raise CommandError("'%s' is not a valid subcommand" %
                                       args.command)
        else:
            self.parser.print_help()

    def main(self,argv):
        parser = self.get_base_parse()
        (options, args) = parser.parse_known_args(argv)
        subcommand_parser= self.get_subcommand_parser()
        self.parser = subcommand_parser
        # if inpute "pyexcel help" or "pyexcel"
        # output the subcommand_parser's help 
        # that function __doc__
        if options.help or not argv:
            self.do_help(options)
            return 0
        args= subcommand_parser.parse_args(argv)
        if args.func == self.do_help:
            self.do_help(args)
            return 0
        args.func(args)


def main(argv=None):
    if argv is None:
        argv = sys.argv[1:]
    shellmain().main(argv)

