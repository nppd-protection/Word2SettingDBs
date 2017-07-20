"""
GUI interface to the make_line_relay_templates program to generate
application-specific templates on the fly rather than pre-generated.
"""

import sys
if sys.version_info[0] < 3:
    import Tkinter as tk
    import ttk
    import Tkconstants
    import tkFileDialog as filedialog
    import tkMessageBox as messagebox
else:
    import tkinter as tk
    import tkinter.ttk as ttk
    import tkinter.constants as Tkconstants
    import tkinter.filedialog as filedialog
    import tkinter.messagebox as messagebox

# Set up a logger so any errors can go to file to facilitate debugging
import logging
from logging.config import dictConfig

import glob
import re
import os.path
import shutil
from docx import Document
from make_line_relay_templates import make_line_relay_templates
from make_line_relay_templates import std_filenames


def get_default_template():
    """ Find default template based on current standard NPPD filesystem
        structure. This is a tradeoff between making the program work easily
        for our users vs having to update the program code when the
        filing system changes. Returns a tuple with the default folder and
        default template name. The returned file will actually exist. If not
        prospective template can be found, then the tuple (None, None) is
        returned.
    """
    default_dir = r'T:\T&DElectronicFiling\ProtCntrl\P&C Procedures\Design ' \
                  r'Standards\Protection Design Standards'
    fn_pattern = 'Settings Dual SEL-421 Line Relay master Standard Rev*.docx'
    fn_list = glob.glob(os.path.join(default_dir, fn_pattern))
    if len(fn_list) > 0:
        return os.path.abspath(default_dir), os.path.basename(fn_list[-1])
    else:
        return None, None

class LogDisplay(tk.LabelFrame):
    """A simple 'console' to place at the bottom of a Tkinter window """
    def __init__(self, root, **options):
        tk.LabelFrame.__init__(self, root, **options)

        "Console Text space"
        scrollbar = tk.Scrollbar(self)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.console = tk.Text(self, height=10, yscrollcommand=scrollbar.set)
        self.console.pack(fill=tk.BOTH)
        scrollbar.config(command=self.console.yview)

class LoggingToGUI(logging.Handler):
    """ Used to redirect logging output to the widget passed in parameters """
    def __init__(self, console=None):
        logging.Handler.__init__(self)

        self.console = console

    def emit(self, message): # Overwrites the default handler's emit method
        formattedMessage = self.format(message) + '\n'

        # Disabling states so no user can write in it
        self.console.configure(state=tk.NORMAL)
        self.console.insert(tk.END, formattedMessage) #Inserting the logger message in the widget
        self.console.configure(state=tk.DISABLED)
        self.console.see(tk.END)
        self.console.update_idletasks()

class TkMakeCalcsGUI(ttk.Frame):

    def __init__(self, root, logger):
        self.root = root
        self.logger = logger

        default_dir, default_fn = get_default_template()
        self.template_opt = options = {}
        options['filetypes'] = [('Word Document', '.docx')]
        options['parent'] = root

        self.saveas_opt = options = {}
        options['filetypes'] = [('Word Document', '.docx')]
        options['initialdir'] = r'T:\T&DElectronicFiling\ProtCntrl\\'
        options['parent'] = root

        root.wm_title('Create calculation template from master')
        ttk.Frame.__init__(self, root)
        button_opt = {'anchor': 'w', 'expand': 1,
                      'padx': 5, 'pady': 5}
        _ = ttk.Frame(self)
        _.pack(**button_opt)
        ttk.Button(_, text='Select Template',
                       command=self.askopenfilename).pack(
            side=Tkconstants.LEFT, **button_opt)

        self.documentParam = tk.StringVar()
        if default_dir is not None and default_fn is not None:
            self.documentParam.set(os.path.join(default_dir, default_fn))
        ttk.Entry(_, textvariable=self.documentParam, width=160).pack(
            side=Tkconstants.LEFT, fill=Tkconstants.BOTH, **button_opt)

        self.std = tk.StringVar()
        self.std.set(sorted(std_filenames)[0])
        for k in sorted(std_filenames):
            v = std_filenames[k]
            ttk.Radiobutton(self, text=k + ': ' + v,
                                value=k, variable=self.std).pack(**button_opt)

        ttk.Button(self, text='Create customized output',
                   command=self.doit).pack(**button_opt)

        log_console = LogDisplay(self)
        log_console.pack(fill='both', **button_opt)
        GUI_handler = LoggingToGUI(console=log_console.console)
        GUI_handler.setLevel(logging.DEBUG)
        GUI_handler.setFormatter(logging.Formatter('%(levelname)-8s %('
                                                   'message)s'))
        logger.addHandler(GUI_handler)


    def askopenfilename(self):
        logger = self.logger
        logger.info('Selecting input file...')
        initialdir, initialfile = os.path.split(self.documentParam.get())
        open_file = filedialog.askopenfilename(
            initialdir=initialdir, initialfile=initialfile,
            **self.template_opt)
        if open_file:
            self.documentParam.set()
            logger.info('Selected input file: ' + self.documentParam.get())

    def doit(self):
        logger = self.logger
        try:
            # Prompt for filename
            logger.info('Selecting output file...')
            documentParam = self.documentParam.get()
            std = self.std.get()
            doc_base = re.match('(.*)\.doc[xm]$', documentParam,
                                flags=re.I).group(1)
            file_rev = re.search(' Rev ?[0-9]+$', doc_base, flags=re.I).group(0)
            initialfile = 'Settings ' + std + ' ' + std_filenames[std] + \
                                     file_rev + '.docx'
            save_file = filedialog.asksaveasfilename(
                initialfile=initialfile,
                **self.saveas_opt)
            # Get parameters from GUI again in case they somehow changed
            documentParam = self.documentParam.get()
            std = self.std.get()

            if save_file:
                logger.info('Selected output file: ' + save_file)
                logger.info('Generating output...')
                logger.info('Using input file: ' + documentParam)
                logger.info('Copying template to: %s' % save_file)
                try:
                    shutil.copyfile(documentParam, save_file)
                except IOError as e:
                    logger.error('Program error', exc_info=True)
                    messagebox.showerror('Error!',
                                         'Selected input file does not exist. '
                                         'Please select the correct template '
                                         'file and try again.')
                    return
                document = Document(save_file)
                logger.info('Customizing to standard: %s' % std)
                make_line_relay_templates(document, std)
                logger.info('Saving output file: %s' % save_file)
                document.save(save_file)
                logger.info('DONE.')
                messagebox.showinfo('Success!', 'Customized output was '
                                                'created at %s' % save_file)
                self.root.destroy()
        except (SystemExit, KeyboardInterrupt):
            raise
        except Exception as e:
            logger.error('Program error', exc_info=True)
            messagebox.showerror('Error!', 'An error occurred. See program log '
                                           'for details.')
            raise

def main():
    import os
    import sys

    # By default, log to the same directory the program is run from
    if os.path.exists(os.path.dirname(sys.argv[0])):
        logfile = os.path.join(os.path.dirname(sys.argv[0]),
                               'make_line_relay_gui.log')
    else:
        logfile = 'make_line_relay_gui.log'

    logging_config = {
        'version': 1,
        'formatters': {
            'file': {'format':
                     '%(asctime)s ' + os.environ['USERNAME'] +
                     ' %(levelname)-8s %(message)s'},
            'console': {'format':
                        '%(levelname)-8s %(message)s'}
            },

        'handlers': {
            'file': {'class': 'logging.FileHandler',
                     'filename': logfile,
                     'formatter': 'file',
                     'level': 'INFO'},
            'console': {'class': 'logging.StreamHandler',
                        'formatter': 'console',
                        'level': 'DEBUG'}
            },
        'loggers': {
            'root': {'handlers': ['file', 'console'],
                     'level': 'DEBUG'}
            }
    }

    dictConfig(logging_config)

    logger = logging.getLogger('root')

    try:
        root = tk.Tk()
        TkMakeCalcsGUI(root, logger).pack()
        root.mainloop()

    except (SystemExit, KeyboardInterrupt):
        raise
    except Exception as e:
        logger.error('Program error', exc_info=True)
        raise

if __name__ == '__main__':
    main()
