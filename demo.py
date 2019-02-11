import VoiceAnalyzer
import sys, getopt

def main(argv):
    inputDir = ''
    outputDir = ''
    try:
        opts, args = getopt.getopt(argv, "hi:o:", ["idir=","odir="])
    except getopt.GetoptError:
        print("for correct usage use: \n")
        print( "demo.py -i <inputDirectory> -o <outputDirectory>")
        sys.exit(2)
    for opt, arg in opts:
        if opt == "-h":
            print("demo.py -i <inputDirectory> -o <outputDirectory>")
            exit()
        elif opt in ("-i", "--idir"):
            inputDir = arg
        elif opt in ("-o", "--odir"):
            outputDir = arg
        else:
            print("for correct usage use: \n")
            print("demo.py -i <inputDirectory> -o <outputDirectory>")
            sys.exit(2)

    if(inputDir == ''):
        print("for correct usage use: \n")
        print("demo.py -i <inputDirectory> -o <outputDirectory> \n inputDirectory is mandatory")
        sys.exit(2)
    if outputDir == '':
        VoiceAnalyzer.processDirectory(inputDir)
    else:
        VoiceAnalyzer.processDirectory(inputDir, outputDir)


if __name__ == "__main__":
    main(sys.argv[1:])
