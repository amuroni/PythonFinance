from v2.Model import Model

Model = Model()  # declare Financial Model variable

irr = Model.summary.range("H22").value  # select a cell, the IRR in this case
print("{:.2%}".format(irr))
