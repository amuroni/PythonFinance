from v2.Model import Model

Model = Model()  # declare Financial Model variable
Model.input()  # set input assumption in the model
Model.check()
irr = Model.summary.range("H22").value  # select a cell, the IRR in this case
print("{:.2%}".format(irr))  # print result in % format
