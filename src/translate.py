import os
import sys
import csv
import argparse


__VERSION__ = "0.0.1"

def main(inputFilePath, outputFilepath):
    columnNames=["Name","Number","Description","OverrideColor","OverrideStyle","OverrideWeight","OverrideMaterial","ByLevelColor","ByLevelStyle","ByLevelWeight","ByLevelMaterial","DisplayPriority","Transparency","GlobalDisplay","GlobalFreeze","ElementAccess","Plot","OverrideStyleScale","OverrideStyleOriginWidth","OverrideStyleEndWidth","ByLevelStyleScale","ByLevelStyleOriginWidth","ByLevelStyleEndWidth"]
    columnDefault=["", 0, "", 0, 0, 0, None, 0, 0, 0, None, 0, 0.000000, 1, 0, 0, 1, "", "", "", "", "", ""]
    columnMapping=[
        { "ms_nr": 0, "ms_name": "Name", "name": "", "nr": 1 },
        { "ms_nr": 7,  "ms_name": "ByLevelColor", "name": "Farbe", "nr": 2 },
        { "ms_nr": 8,  "ms_name": "ByLevelStyle", "name": "Strichart", "nr": 4 },
        { "ms_nr": 9, "ms_name": "ByLevelWeight", "name": "Strichstaerke", "nr": 3 },
        { "ms_nr": 11, "ms_name": "DisplayPriority", "name": "Prioritaet", "nr": 5 },
        { "ms_nr": 12, "ms_name": "Transparency", "name": "Transparenz", "nr": 6 }
    ]

    with open(inputFilePath, mode='r') as csvfile:
        with open(outputFilepath, mode='w', newline='') as outfile:
            layerDefinition = csv.reader(csvfile, delimiter=';')
            output = csv.writer(outfile, delimiter=',')

            output.writerow(["%SECTION","Levels"])
            output.writerow("")
            output.writerow(columnNames)

            lastPrefix = ""
            for row in layerDefinition:
                if row[0]:
                    lastPrefix = row[0]
                newrow = columnDefault
                for m in columnMapping:
                    val = row[m["nr"]]
                    if m["ms_nr"] == 0:
                        val = f"{lastPrefix}{val}"
                    newrow[m["ms_nr"]] = val
                output.writerow(newrow)

            output.writerow("")
            output.writerow(["%SECTION","level-filter"])
            output.writerow("")
            output.writerow(["filter-name","filter-parent","level-group","filter-compose","level-name","level-code","level-description","level-color","level-style","level-weight","level-material","level-element-color","level-element-style","level-element-weight","level-element-material","level-display","level-frozen","level-priority","level-plot","level-used","level-transparency","level-element-access"])

parser = argparse.ArgumentParser()
parser.add_argument("--output", help="Ausgabedatei")
parser.add_argument("--input", help="Eingabedatei")
parser.add_argument("--version", help="Version anzeigen", action="store_true")
args = parser.parse_args()

if args.version:
    print(f"version: {__VERSION__}")
    sys.exit(0)

if not args.input:
    print("Bitte input Datei angeben")
    sys.exit(0)

if not args.output:
    print("Bitte output Datei angeben")
    sys.exit(0)

main(args.input, args.output)
