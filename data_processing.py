import pandas as pd
import numpy as np
dataset_path = "dataset.xlsx"

dataset = pd.read_excel(dataset_path, sheet_name = 'Sheet1')

sentences = dataset["SENTENCE"]

main_intents = dataset["MAIN"]

qas = dataset["QA"]


# result_intents is dictionary {"intent1":[(sentence1, qa1), (sentence2, qa2)], "intent2":[(sentence3, qa3), (sentence4, qa4), (sentence5, qa5), ...]}

result_intents = {}


# fill result_intents according to main_intent
count_intent = {}
for idx, main_intent in enumerate(main_intents):

    if main_intent not in result_intents.keys():

        result_intents[main_intent] = [(sentences[idx], qas[idx])]
        count_intent[main_intent] = 0

    else:

        result_intents[main_intent].append((sentences[idx], qas[idx]))
        count_intent[main_intent] = count_intent[main_intent] + 1
x = count_intent.values()
#sorted(count_intent)

print(sorted(count_intent.items(), key=lambda k: k[1], reverse=True))

#count_intent.s
#print(count_intent.sort())

