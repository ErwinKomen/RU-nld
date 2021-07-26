"""
A more generic textsearch processor
THis version created by Erwin R. Komen
Date: 14/jul/2021
"""
import re
import utils


class TextSearch(object):
    wordlist = []

    def __init__(self, wordlist, *args, **kwargs):
        self.wordlist = []
        # Copy the wordlist
        for word in wordlist:
            self.wordlist.append(word.lower())
        # Perform the generic object init stuff
        response = super(TextSearch, self).__init__(*args, **kwargs)

        # Other initialisations
        self.errHandle = utils.ErrHandle()
        self.reLineEnd = re.compile(r"[\.]")
        self.reNalpha = re.compile(r"[^\w\(\)\:\;\!\?]")

        # Return the response that is needed
        return response

    def exists(self, text_part, bShow = False):
        """Check if [text_part] contains anything from the list [wordlist]"""

        method = "word_by_word"
        method = "sentences"

        iCount = 0
        try:
            text_part = text_part.strip()
            if text_part != "":
                # Change the newlines to spaces
                text_part = text_part.replace("\n", " ")
                text_part = text_part.replace("\r", "")
                # Divide the line into parts divided by sentence breakers [.], [?], [!]
                lLine = re.sub(self.reLineEnd, "\n", text_part).split(sep="\n")

                # Walk all lines
                for sText in lLine:
                    # Trim it
                    sText = sText.strip().strip("\n").strip("\r")

                    # double check if it is empty
                    if sText != "":
                        if method == "word_by_word":
                            # Tokenize the text into words on the basis of spaces, stripping off metadata
                            wList = re.sub(self.reNalpha, " ", sText).split()
                            # Walk the list of words
                            for wrd in wList:
                                # Make sure the word is lower-case
                                wrd = wrd.lower()
                                # Check if the word matches with anything in the word list
                                if wrd in self.wordlist:
                                    iCount += 1
                                    break
                        else:
                            for chunk in self.wordlist:
                                # if chunk in sText:
                                match_chunk = r'.*\b{}\b.*'.format(re.escape(chunk))
                                if re.match(match_chunk, sText, re.IGNORECASE):
                                    iCount += 1
                                    # DEBUGGING: show the hit
                                    if bShow:
                                        self.errHandle.Status("Hit: [{}] in: {}".format(chunk, sText))
                                    break
                                
            # The result is a count higher than 0
            return (iCount > 0)
        except:
            # act
            self.errHandle.DoError("TextSearch/exists exception")
            return False

