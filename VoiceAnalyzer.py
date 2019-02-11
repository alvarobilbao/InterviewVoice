import subprocess
import speech_recognition
import base64
from speech_recognition import AudioData, RequestError, PortableNamedTemporaryFile, UnknownValueError
from urllib.error import URLError
import googleapiclient
import os
import xlsxwriter
from rake_nltk import Rake
import json
#CHANGE LOGIC TO USE GOOGLE AUTH INSTEAD OG OAUTH2
def my_recognize_google_cloud(self, audio_data, credentials_json=None, language="en-US", preferred_phrases=None,
                           show_all=False):
    """
    Performs speech recognition on ``audio_data`` (an ``AudioData`` instance), using the Google Cloud Speech API.
    This function requires a Google Cloud Platform account; see the `Google Cloud Speech API Quickstart <https://cloud.google.com/speech/docs/getting-started>`__ for details and instructions. Basically, create a project, enable billing for the project, enable the Google Cloud Speech API for the project, and set up Service Account Key credentials for the project. The result is a JSON file containing the API credentials.
    The path to this JSON file is specified by ``credentials_json``. If not specified, the library will try to automatically `find the default API credentials JSON file <https://developers.google.com/identity/protocols/application-default-credentials> (remember to define GOOGLE_APPLICATION_CREDENTIALS environment variable)`__.
    The recognition language is determined by ``language``, which is a BCP-47 language tag like ``"en-US"`` (US English). A list of supported language tags can be found in the `Google Cloud Speech API documentation <https://cloud.google.com/speech/docs/languages>`__.
    If ``preferred_phrases`` is an iterable of phrase strings, those given phrases will be more likely to be recognized over similar-sounding alternatives. This is useful for things like keyword/command recognition or adding new phrases that aren't in Google's vocabulary. Note that the API imposes certain `restrictions on the list of phrase strings <https://cloud.google.com/speech/limits#content>`__.
    Returns the most likely transcription if ``show_all`` is False (the default). Otherwise, returns the raw API response as a JSON dictionary.
    Raises a ``speech_recognition.UnknownValueError`` exception if the speech is unintelligible. Raises a ``speech_recognition.RequestError`` exception if the speech recognition operation failed, if the credentials aren't valid, or if there is no Internet connection.
    """
    assert isinstance(audio_data, AudioData), "``audio_data`` must be audio data"
    assert isinstance(language, str), "``language`` must be a string"
    assert preferred_phrases is None or all(
        isinstance(preferred_phrases, (type(""), type(u""))) for preferred_phrases in
        preferred_phrases), "``preferred_phrases`` must be a list of strings"

    # See https://cloud.google.com/speech/reference/rest/v1/RecognitionConfig
    flac_data = audio_data.get_flac_data(
        convert_rate=None if 8000 <= audio_data.sample_rate <= 48000 else max(8000, min(audio_data.sample_rate, 48000)),
        # audio sample rate must be between 8 kHz and 48 kHz inclusive - clamp sample rate into this range
        convert_width=2  # audio samples must be 16-bit
    )

    try:
        #from oauth2client.client import GoogleCredentials
        from googleapiclient.discovery import build
        import googleapiclient.errors
        import google.auth
        from google.oauth2 import service_account
        # cannot simply use 'http = httplib2.Http(timeout=self.operation_timeout)'
        # because discovery.build() says 'Arguments http and credentials are mutually exclusive'
        import socket
        import googleapiclient.http
        if self.operation_timeout and socket.getdefaulttimeout() is None:
            # override constant (used by googleapiclient.http.build_http())
            googleapiclient.http.DEFAULT_HTTP_TIMEOUT_SEC = self.operation_timeout

        if credentials_json is None:
            api_credentials = google.auth.default()
        else:
            api_credentials = service_account.Credentials.from_service_account_file(credentials_json)
            # the credentials can only be read from a file, so we'll make a temp file and write in the contents to work around that
            #with PortableNamedTemporaryFile("w") as f:
            #    f.write(credentials_json)
            #    f.flush()
            #    api_credentials = GoogleCredentials.from_stream(f.name)

        speech_service = build("speech", "v1", credentials=api_credentials, cache_discovery=False)
    except ImportError:
        raise RequestError(
            "missing google-api-python-client module: ensure that google-api-python-client is set up correctly.")

    speech_config = {"encoding": "FLAC", "sampleRateHertz": audio_data.sample_rate, "languageCode": language}
    if preferred_phrases is not None:
        speech_config["speechContexts"] = [{"phrases": preferred_phrases}]
    if show_all:
        speech_config["enableWordTimeOffsets"] = True  # some useful extra options for when we want all the output
    request = speech_service.speech().recognize(
        body={"audio": {"content": base64.b64encode(flac_data).decode("utf8")}, "config": speech_config})

    try:
        response = request.execute()
    except googleapiclient.errors.HttpError as e:
        raise RequestError(e)
    except URLError as e:
        raise RequestError("recognition connection failed: {0}".format(e.reason))

    if show_all: return response
    if "results" not in response or len(response["results"]) == 0: raise UnknownValueError()
    transcript = ""
    averageConfidence = 0
    numberOfTranscripts = 0
    for result in response["results"]:
        transcript += result["alternatives"][0]["transcript"].strip() + " "
        averageConfidence += result["alternatives"][0]["confidence"]
        numberOfTranscripts += 1

    averageConfidence /= numberOfTranscripts
    return {
        'transcript': transcript,
        'confidence': averageConfidence
    }

def get_google_nlp_service(credential_json_path=None):
    """
        Sets the credentials to get the google Language API service, that will enable to call the entities and the sentiments recognition services
        :param credential_json_path : String referring to the JSON file with the google cloud credentials
    """

    from googleapiclient.discovery import build
    import google.auth
    from google.oauth2 import service_account

    if credential_json_path is None:
        api_credentials = google.auth.default()
    else:
        api_credentials = service_account.Credentials.from_service_account_file(credential_json_path)

    nlp_service = build("language", "v1", credentials=api_credentials, cache_discovery=False)

    return nlp_service

def callNLPService(text):
    """
    Call to the google language API to recognize entities and sentiments on a text
    :param text: String the text to be analyzed
    :return: Array of entities according to google API
    """
    google_cloud_credentials = "./assets/Interview_Voice_google_cloud_key.json"
    nlp_service = get_google_nlp_service(google_cloud_credentials)
    client = nlp_service.documents()
    request1 = client.analyzeEntitySentiment(body={
        "document": {
            "type": "PLAIN_TEXT",
            "content": text,
            "language": "en_IN"
        }
    })
    try:
        response = request1.execute()
    except googleapiclient.errors.HttpError as e:
        raise RequestError(e)
    except URLError as e:
        raise RequestError("recognition connection failed: {0}".format(e.reason))
    entities = response["entities"]
    return entities

def writeEntitiesDocument(outputDirPath, basename, entities):
    """
    Generate the txt file containing all the entities and the details related to them
    :param outputDirPath: String refering to the directory where the txt file will be placed
    :param basename: the name of the video file that was analyzed
    :param entities: list of the entities
    :return: Array of the entities name property
    """
    f = open(outputDirPath + "/" + basename + "NLPResult.txt", "w+")
    words = []
    for entity in entities:
        f.write('=' * 20)
        f.write('\n')
        f.write(u'{:<16}: {}\n'.format('name', entity["name"]))
        words.append(entity["name"])
        f.write(u'{:<16}: {}\n'.format('sentiment', entity["sentiment"]))
        f.write(u'{:<16}: {}\n'.format('type', entity["type"]))
        f.write(u'{:<16}: {}\n'.format('metadata', entity["metadata"]))
        f.write(u'{:<16}: {}\n'.format('salience', entity["salience"]))
        f.write(u'{:<16}: {}\n'.format('wikipedia_url',
                                        entity["metadata"].get('wikipedia_url', '-')))
    return words

def analyzeSentiments(text):
    """
    Calls the analyse sentiment service of the google language API
    :param text: String that will be analysed by the service
    :return: Object containing the sentiment towards the text that was analyzed
    """
    google_cloud_credentials = "./assets/Interview_Voice_google_cloud_key.json"
    nlp_service = get_google_nlp_service(google_cloud_credentials)
    client = nlp_service.documents()
    request2 = client.analyzeSentiment(body={
        "document": {
            "type": "PLAIN_TEXT",
            "content": text,
            "language": "en_IN"
        }
    })

    try:
        response = request2.execute()
    except googleapiclient.errors.HttpError as e:
        raise RequestError(e)
    except URLError as e:
        raise RequestError("recognition connection failed: {0}".format(e.reason))
    sentiment = response["documentSentiment"]
    return sentiment

def processVideo(pathToFile, outputDirPath="./"):
    """
    Main function of the API, it process the video and gathers all the information from it (relevant phrases, entities)
    :param pathToFile: String, the file path of the video to be analyzed
    :param outputDirPath: String the directory path where all the files that are generated will be placed
    :return: Object containing the Rake results, the entities result, the setiment analysis and the confidence of the speech to text translation
    """
    basename = os.path.basename(pathToFile)
    basename = os.path.splitext(basename)[0]
    print("pathToFile " + pathToFile)
    wavFilePath = (outputDirPath + "/" + basename + ".wav").replace(" ", "").replace("&", "And")

    command = "ffmpeg -y -i \"" + pathToFile + "\" -ab 160k -ac 2 -ar 44100 -vn " + wavFilePath
    print("waveFile path = " + wavFilePath)
    subprocess.run(command, shell=True)

    r = speech_recognition.Recognizer()
    with speech_recognition.AudioFile(wavFilePath) as source:
        audio = r.record(source)

    try:
        r.recognize_google_cloud = my_recognize_google_cloud
        google_cloud_credentials = "./assets/Interview_Voice_google_cloud_key.json"
        recognizedAudioResult = r.recognize_google_cloud(r, audio, credentials_json=google_cloud_credentials, language="en_IN")
        recognizedTranscript = recognizedAudioResult['transcript']
        confidence = recognizedAudioResult['confidence']
        print("Writting speech recognition to File: \n")
        f = open(outputDirPath + "/" + basename + "Text.txt", "w+")
        f.write(recognizedTranscript)
        entities = callNLPService(recognizedTranscript)
        words = writeEntitiesDocument(outputDirPath, basename, entities)
        phrases = extract_phrases(recognizedTranscript)
        writeRakeResults(outputDirPath, basename, phrases)

        sentiment = analyzeSentiments(recognizedTranscript)
        output = {
            "sentiments": sentiment,
            "words": words,
            "confidence": confidence,
            "rakePhrases": phrases
        }
        return output

    except speech_recognition.UnknownValueError:
        print("Oops! Didn't catch that")
    except speech_recognition.RequestError as e:
        print("Uh oh! Couldn't request results from Google Cloud Speech Recognition service; {0}".format(e))

def processDirectory(pathToDirectory, outputDirPath="./"):
    """
    Call the process video function to analyse all th videos on a directory
    :param pathToDirectory: String referring to the directory location
    :param outputDirPath: String the directory where to place the results
    :return: void
    """
    workbook = xlsxwriter.Workbook('./assets/Sentiments.xlsx')
    worksheet = workbook.add_worksheet()
    worksheet.write(0, 0, "file name")
    worksheet.write(0, 1, "magnitude")
    worksheet.write(0, 2, "score")
    worksheet.write(0, 3, "translation to Text confidence")
    row = 1
    for dirName, subdirList, fileList in os.walk(pathToDirectory):
        print("dirname: " + dirName.replace("\\", "/"))
        words = {}
        rakePhrases = {}
        for fileName in fileList:
            if fileName.endswith('.mp4'):
                print("filename = " + fileName)
                videoOutput = processVideo(dirName.replace("\\", "/") + "/" + fileName, dirName.replace("\\", "/"))
                worksheet.write(row, 0, dirName.replace("\\", "/") + "/" + fileName)
                if videoOutput is not None and videoOutput["sentiments"] is not None:
                    print(videoOutput["sentiments"])
                    worksheet.write(row, 1, videoOutput["sentiments"]["magnitude"])
                    worksheet.write(row, 2, videoOutput["sentiments"]["score"])
                    worksheet.write(row, 3, videoOutput["confidence"])
                    row += 1
                if videoOutput is not None and videoOutput["words"] is not None:
                    for word in videoOutput["words"]:
                        if word in words:
                            words[word] += 1
                        else:
                            words[word] = 1
                if videoOutput is not None and videoOutput["rakePhrases"] is not None:
                    for phrase in videoOutput["rakePhrases"]:
                        if phrase in rakePhrases:
                            rakePhrases[phrase[1]] += 1
                        else:
                            rakePhrases[phrase[1]] = 1
        wordsWorkbook = xlsxwriter.Workbook(dirName.replace("\\", "/") + "/recognizedWordsUsage.xlsx")
        wordsWorksheet = wordsWorkbook.add_worksheet()
        wordsWorksheet.write(0, 0, "word")
        wordsWorksheet.write(0, 1, "count")
        wordIndex = 1
        for key, value in words.items():
            wordsWorksheet.write(wordIndex, 0, key)
            wordsWorksheet.write(wordIndex, 1, value)
            wordIndex += 1
        wordsWorkbook.close()

        rakePhrasesWorkbook = xlsxwriter.Workbook(dirName.replace("\\", "/") + "/RakePhrasesUsageByQuestion.xlsx")
        rakePhrasesWorksheet = rakePhrasesWorkbook.add_worksheet()
        rakePhrasesWorksheet.write(0, 0, "word")
        rakePhrasesWorksheet.write(0, 1, "count")
        rakePhrasesIndex = 1
        for key, value in rakePhrases.items():
            rakePhrasesWorksheet.write(rakePhrasesIndex, 0, key)
            rakePhrasesWorksheet.write(rakePhrasesIndex, 1, value)
            rakePhrasesIndex += 1
        rakePhrasesWorkbook.close()
    workbook.close()

def extract_phrases(text):
    """
    Calls the RAKE API to extract the relevant phrases of the given text
    :param text: String, the text to be analyzed
    :return: Array containing the phrases and their scores
    """
    extractor = Rake()
    extractor.extract_keywords_from_text(text)
    return extractor.get_ranked_phrases_with_scores()

def writeRakeResults(outputDirPath, basename, phrases):
    """
    Function to save the RAKE analysis results to an spreadsheet
    :param outputDirPath: String Location where to create this file
    :param basename: String - name of the analysed video
    :param phrases: Array - containing all the phrases and their scores
    :return: void
    """
    workbook = xlsxwriter.Workbook(outputDirPath + "/" + basename + 'RAKE.xlsx')
    worksheet = workbook.add_worksheet()
    worksheet.write(0, 0, "phrase")
    worksheet.write(0, 1, "score")
    row = 1
    for phrase in phrases:
        worksheet.write(row, 0, phrase[1])
        worksheet.write(row, 1, phrase[0])
        row += 1
    workbook.close()


if __name__ == '__main__':
    pathToFile = "./assets/chrisfansi96.mp4"
    #pathToDir = "./assets/Why_are_you_interested_in_persuing_a_master_degree_in"
    pathToDir = "./assets/ByQuestion"
    outputPath = "./assets/Why_are_you_interested_in_persuing_a_master_degree_in/results"
    #processVideo(pathToFile)
    processDirectory(pathToDir, outputPath)
