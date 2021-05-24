import os


def string_search():
    """
    Takes in a path, filename starting string, filename ending string, and a search string.
    Searching the directory for files matching the filename parameters and returns
    a list of filenames that contain the search string.

    Note: user input is case insensitive.
    """
    dir_path = input('What is the directory path containing the files you want to search?: ')
    start_of_filename = input('What is the beginning of the filename(s)? (Leave blank if not needed): ').lower()
    end_of_filename = input('What is the end of the filename(s)? (Leave blank if not needed): ').lower()
    search_string = input('What string are you searching for in these files?: ').lower()

    filenames = []

    for root, dirs, files in os.walk(dir_path):
        for file in files:
            if file.lower().startswith(start_of_filename) and file.lower().endswith(end_of_filename):
                full_path = dir_path + '\\' + file
                f = open(full_path, 'rt')
                content = f.readlines()
                for line in content:
                    if line.lower().find(search_string) != -1:
                        filenames.append(file)
                f.close()

    filenames.sort()

    if not len(filenames):
        print('\nYour search did not return any results.\n')
    else:
        print('\nHere is the list of filenames containing your string: \n')

    for name in filenames:
        print(name)


string_search()
