from sys import path
import os

path.append(os.path.join(os.path.dirname(os.path.dirname(__file__)), 'packages'))

import extra.iota

print(extra.iota.FunI())
