import pandas as pd
import glob
import re



def vocab_files_dir(input_dir, stories_filename, stories_colname):
    '''
    input_dir = vocab file where the first column must contain the kanji-containing Japanese vocab.
    stories_filename = filename of our 3-column kanji stories file.
    stories_colname = column name for the additional column in which stories will be aded.
    '''
    input_files = sorted(glob.glob(input_dir + '/*.xlsx'))
    main_vocab_story_dict = {}
    for i_0 in input_files:
        print("Processing " + i_0)
        # Read the vocabulary files and define columns:
        flashcards_file_df = pd.read_excel(i_0)
        flashcards_file_dict = pd.DataFrame.to_dict(flashcards_file_df, orient='list')
        vocab_kanji_key = list(flashcards_file_dict.keys())[0]
        # Import Excel file of Kodansha stories:
        stories_df = pd.read_excel(stories_filename, header=None)
        stories_df.columns = ["kanji", "meaning", "story"]
        stories_file_dict = pd.DataFrame.to_dict(stories_df, orient='list')
        # Scan each word in the vocab dataframe 'kanji' column for kanji that appear in the stories file:
        parent_vocab_scan_kanji_list = []
        for i_1 in flashcards_file_dict[vocab_kanji_key]:
            child_vocab_scan_kanji_list = []
            for i_2 in i_1:
                for i_3 in stories_file_dict["kanji"]:
                    if i_2 == i_3:
                        if i_3 not in child_vocab_scan_kanji_list:
                            child_vocab_scan_kanji_list.append(i_3)
            parent_vocab_scan_kanji_list.append(child_vocab_scan_kanji_list)
        # Create structured nested lists containing relevant kanji information and colour formatting:
        parent_story_list = []
        for i_1 in parent_vocab_scan_kanji_list:
            child_vocab_scan_kanji_list = []
            for i_2 in i_1:
                for num, i_3 in enumerate(stories_file_dict["kanji"]):
                    gc_list = []
                    if i_2 == i_3:
                        # Add a yellow colour tag to make the kanji more readable:
                        gc_list.append("<color yellow>" + stories_file_dict["kanji"][num] + "</color>")
                        gc_list.append("<color yellow>" + stories_file_dict["meaning"][num] + "</color>")
                        gc_list.append(stories_file_dict["story"][num])
                        child_vocab_scan_kanji_list.append(gc_list)
            parent_story_list.append(child_vocab_scan_kanji_list)
        # Create a single list of strings with Flashcards Deluxe formatting for each vocab:
        story_string_list = []
        for i_1 in parent_story_list:
            if i_1 == []:
                story_string = ""
            else:
                child_story_list = []
                for i_2 in i_1:
                    joined = ("|").join(i_2)
                    child_story_list.append(joined)
                story_string = ("||").join(child_story_list)
            story_string_list.append(story_string)
        # Add our story string list as another value:
        flashcards_file_dict[stories_colname] = story_string_list
        main_vocab_story_dict[i_0] = flashcards_file_dict
    # Return the main dictionary:
    return main_vocab_story_dict


## ---- RUN SCRIPT FROM HERE ---- ##

# Create the dictionary containing vocab data and kanji stories for all vocab files in the input directory:
input_dir = "input_dir"
output_dir = "output_dir"
vocab_story_dict = vocab_files_dir(input_dir = input_dir,
                                   stories_filename = "2021.02.18_all_kodansha_kanji_course.xlsx",
                                   stories_colname = "Text 4")

# Export the Excel files from the dictionary we created into the output directory:
for i_0 in vocab_story_dict.keys():
    output_filename = re.sub(input_dir, output_dir, i_0)
    single_df = pd.DataFrame(vocab_story_dict[i_0])
    print("Exporting " + output_filename)
    single_df.to_excel(output_filename, sheet_name='Sheet1', index=False)
