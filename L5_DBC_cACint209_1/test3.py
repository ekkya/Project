import os
from neuron import h
import numpy
import sys


h.load_file("morphology.hoc")
# Load biophysics
h.load_file("biophysics.hoc")
# Load main cell template
h.load_file("template.hoc")
print
