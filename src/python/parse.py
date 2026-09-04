import json

with open("metricTargets.json", "r") as f:
    metricTarget = json.load(f)

metricTarget_tags = metricTarget["tags"]


def parseMetricFields(
    conceptID: str, metricDP: dict, topicalM: str | None
) -> dict | None:
    metricFields: dict = {}

    if conceptID in [
        "ESRS26_GDR-M_02",
        "ESRS26_GDR-M_03",
        "ESRS26_GDR-M_04",
        "ESRS26_GDR-M_05",
    ]:
        metricFields[conceptID] = {
            "name": metricDP["taxonomyFields"]["name"],
            "qname": metricDP["taxonomyFields"]["qname"],
            "taggedText": metricDP["text"],
        }

    if conceptID == "ESRS26_GDR-M_IroToMetric":
        metricFields[conceptID] = metricDP["text"]

    if topicalM:
        # metric is a topical metric so should not have values and link to target
        metricFields[conceptID] = topicalM
    else:
        if conceptID.startswith("ESRS26_GDR-MetricValue"):
            metricFields[conceptID] = metricDP["value"]

        if conceptID == "ESRS26_GDR-M_TargetToMetric":
            metricFields[conceptID] = metricDP["text"]

    return metricFields or None


def parseMetrics(tags: dict):
    metricIdentifiers = set()
    nameDict: dict = {}
    metricInfos: dict = {}

    for tagID, fields in tags.items():
        topical_metric = None

        conceptID: str = fields["taxonomyFields"]["conceptId"]
        if not conceptID.startswith("ESRS26_GDR-M"):
            continue

        try:
            metricId = fields["dimensions"]["dim10"]["member"]
        except KeyError:
            print(f"Tag: '{tagID}' do not have the metric identifier as dimension.")

        if conceptID == "ESRS26_GDR-M_TopicalMetric":
            valueS = fields["value"]
            if isinstance(valueS, str):
                topical_metric = valueS
            else:
                topical_metric = "".join(str(e) for e in valueS)

        if metricInfos.get(metricId) is None:
            metricInfos[metricId] = []
        else:
            info = parseMetricFields(conceptID, fields, topical_metric)
            if info is not None:
                metricInfos[metricId].append(info)

        if conceptID == "ESRS26_GDR-MetricName":
            nameMetric = fields["text"]
            nameDict[metricId] = nameMetric

        metricIdentifiers.add(metricId)

    print("\nThe following metrics are stored in the data:")
    for i in metricIdentifiers:
        if nameDict.get(i):
            print(f"Identifier: {i}   Name: {nameDict[i]}")
        else:
            print(f"Identifier: {i}")

    identifierInput = input("\nInsert metric identifier you want to analyze: ")
    if identifierInput not in metricIdentifiers:
        print("Identifier inputted is not valid.\n")

        previous = identifierInput
        for i in metricIdentifiers:
            if identifierInput[:4] == i[:4]:
                answer = input(f"Did you mean '{i}'? Type 'y' or 'n': ")
                if answer == "y":
                    identifierInput = i
                    break
                else:
                    continue

        if identifierInput == previous:
            return

    metricInfo = metricInfos[identifierInput]
    for information in metricInfo:
        taxID = next(iter(information))
        info = next(iter(information.values()))
        if isinstance(info, dict):
            print(f"\n   {taxID}:\n")
            for k, v in info.items():
                print(f"      {k}: {v}")
        else:
            print(f"\n   {taxID}: {info}")
    print("\n")
