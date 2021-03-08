import pandas
import copy
import os.path

# Make directory for stories files:
if os.path.isdir("kodansha_vocab_stories"):
    print("Stories directory exists. Do not make.")
else:
    os.mkdir("kodansha_vocab_stories")
    print("Making stories directory.")

# Step 1: Import Excel file of flashcards as pandas table.
n = 1
max_deck_no = 18 # Maximum number of decks to make.
while n <= max_deck_no:
    # Step 1: Read the vocabulary files:
    n_str = str(n).zfill(2)
    flashcards_file = "kodansha_vocab/kodansha_kanji_course_vocab_" + n_str + "_of_18.xlsx"
    new_flashcards_file = "kodansha_vocab_stories/kodansha_kanji_course_vocab_stories_" + n_str + "_of_18.xlsx"
    flashcards_file_df = pandas.read_excel(flashcards_file)
    flashcards_file_df_colnames = ['Text 1', 'Text 2', 'Text 3']
    flashcards_file_df.columns = flashcards_file_df_colnames
    flashcards_kanji_list = flashcards_file_df['Text 1'].tolist()
    # Step 2: Import Excel file of my Kodansha stories:
    stories_filename = "2021.02.18_all_kodansha_kanji_course.xlsx"
    stories_df = pandas.read_excel(stories_filename, header=None)
    stories_df.columns = ["kanji", "meaning", "story"]
    stories_kanji_list = stories_df.kanji.tolist()
    stories_meaning_list = stories_df.meaning.tolist()
    stories_story_list = stories_df.story.tolist()
    filt_stories_story_list = []
    for i_0 in stories_story_list:
        story_var = copy.deepcopy(i_0)
        if isinstance(story_var, float):
            story_var = ""
        filt_stories_story_list.append(story_var)
    stories_story_list = filt_stories_story_list
    # Step 3: Scan each Japanese word in flashcards table, extract a list of kanji for each word:
    parent_vocab_scan_kanji_list = []
    for i_0 in flashcards_kanji_list:
        vocab_scan = copy.deepcopy(i_0)
        child_vocab_scan_kanji_list = []
        for i_1 in vocab_scan:
            for i_2 in stories_kanji_list:
                if i_1 == i_2:
                    if i_2 not in child_vocab_scan_kanji_list:
                        child_vocab_scan_kanji_list.append(i_2)
        parent_vocab_scan_kanji_list.append(child_vocab_scan_kanji_list)
    parent_story_list = []
    for i_0 in parent_vocab_scan_kanji_list:
        vocab_scan = copy.deepcopy(i_0)
        child_vocab_scan_kanji_list = []
        for i_1 in vocab_scan:
            for i_2 in range(len(stories_kanji_list)):
                story_kanji = copy.deepcopy(stories_kanji_list[i_2])
                gc_list = []
                if i_1 == story_kanji:
                    # Add a yellow colour tag to make the kanji more readable:
                    gc_list.append("<color yellow>" + story_kanji + "</color>")
                    gc_list.append("<color yellow>" + stories_meaning_list[i_2] + "</color>")
                    gc_list.append(stories_story_list[i_2])
                    child_vocab_scan_kanji_list.append(gc_list)
        parent_story_list.append(child_vocab_scan_kanji_list)
    story_string_list = []
    for i_0 in parent_story_list:
        story_scan = copy.deepcopy(i_0)
        if story_scan == []:
            story_string = ""
        else:
            child_story_list = []
            for i_1 in story_scan:
                joined = ("|").join(i_1)
                child_story_list.append(joined)
            story_string = ("||").join(child_story_list)
        story_string_list.append(story_string)
    # Create the final Excel file:
    flashcards_file_df["Text 4"] = story_string_list
    flashcards_file_df_colnames = ['Text 1', 'Text 2', 'Text 3', 'Text 4']
    flashcards_file_df.columns = flashcards_file_df_colnames
    flashcards_file_df.to_excel(new_flashcards_file, sheet_name='Sheet1', index=False)
    n += 1
