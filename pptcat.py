"""
A utility to help find slides in large piles of PowerPoint docs.

Roadmap

v1:
- db + file system
- search using sqlite tools (e.g. SQLiteExpert), image browser, grep, etc

v2:
- db only with search tool

Todo:

* fix textonly mis-classification cases
* fix so all writing for a ppt file is a single transaction
* add file metadata (author, creation date, ...)
* enable SQLite full text search (https://charlesleifer.com/blog/using-sqlite-full-text-search-with-python/)
* browser based search tool
* run over everything at scion
* free from needing PowerPoint installed
* pkg as exe
* semantic image search? (or at least ordeering by similarity)

"""


import sys
import os
import logging
logging.basicConfig(level=logging.INFO)

SCHEMASQL = """
create table if not exists files(
  rowid integer primary key autoincrement,
  filename text not null unique,
  hash text not null unique,
  nslides integer
);
create table if not exists slides(
  fileid integer,
  islide integer,
  fingerprint text,
  thumb blob,
  hires blob,
  text text,
  textonly integer,
  foreign key (fileid)
    references files (rowid)
    on delete cascade
    on update cascade,
  unique (fileid, islide)
);
"""


def make_temp_dir():
    import tempfile
    return tempfile.mkdtemp()


def extract_slides(fn):
    import comtypes.client

    # surely these must be importable from somewhere
    msoGroup = 6 # msoShapeType Enum
    msoTrue = -1 #!!! wtf??? msoTriState Enum

    def text_from_group(parent):
        text = []
        for child in parent.GroupItems:
            if child.Type==msoGroup:
                text.extend(text_from_group(child))
            else:
                if child.HasTextFrame==msoTrue and child.TextFrame.HasText==msoTrue:
                    text.append(child.TextFrame.TextRange.Text)
        return text


    def contains_types(objs, types=(30,1,2,20,3,27,21,7,8,5,28,24,22,23,9,31,29,10,11,16,12,13,-2,19,26)):
        # default types is (hopefully) anything that isn't text
        # todo: do msoPlaceholder=14 objects have children???
        for obj in objs:
            #print(obj.Type)
            if obj.Type in types:
                return True
            elif obj.Type==msoGroup:
                if contains_types(obj.GroupItems, types=types):
                    return True
        return False


    def render_slide(height):
        from PIL import Image
        fn = os.path.join(tmp_dir, '%i_thumb.png' % islide)
        slide.Export(fn, 'PNG', ScaleWidth=(height*16)//9, ScaleHeight=height)
        img = Image.open(fn)
        img.load()
        return img


    slides = []
    tmp_dir = make_temp_dir()
    logging.debug('using temp dir %s', tmp_dir)

    powerpoint = comtypes.client.CreateObject("Powerpoint.Application")
    powerpoint.Visible = True
    powerpoint.Presentations.Open(fn)
    for islide, slide in enumerate(powerpoint.ActivePresentation.Slides):
        logging.debug('parsing slide %s', islide)
        this = {'filename': os.path.abspath(fn),
                'islide': islide+1, # to match Powerpoint's slide numbering
        }

        # extract text
        text = []
        for shp in slide.Shapes:
            #print(shp.HasTextFrame)
            if shp.HasTextFrame==msoTrue and shp.TextFrame.HasText==msoTrue:
                #import pdb; pdb.set_trace()
                text.append(shp.TextFrame.TextRange.Text)
            if shp.Type==msoGroup:
                text.extend(text_from_group(shp))
        this['text'] = text

        # extract images
        this['thumb'] = render_slide(height=240) # PIL.Image object
        this['hires'] = render_slide(height=1080)

        # figure out if this slide contains anything other than text
        this['textonly'] = not contains_types(slide.Shapes) # default is to look for non text types

        # todo: serialize slide?

        slides.append(this)

    powerpoint.Presentations[1].Close()
    powerpoint.Quit()

    # todo: cleanup tmp dir

    return slides


def db_connect(dbfn='pptcat.db3'):
    import sqlite3
    db = sqlite3.connect(dbfn)
    cur = db.cursor()
    cur.executescript(SCHEMASQL) # safe due to "if not exists"
    db.commit()
    db.row_factory = sqlite3.Row
    return db


def get_files_to_index():
    fns = []
    for arg in sys.argv[1:]:
        if os.path.isfile(arg):
            fns.append(arg)
        elif os.path.isdir(arg):
            #from glob import glob
            #fns += glob(os.path.join(arg, '**/*.ppt[x]?'), recursive=True)
            import re
            regexp = re.compile('^.*\.[pP][pP][tT][xX]?$')
            for root, dirnames, filenames in os.walk(arg):
                fns += [os.path.join(root, fn) for fn in filenames if regexp.match(fn)]
    return fns


def file_checksum(fn):
    import hashlib
    hash_md5 = hashlib.md5()
    with open(fn, "rb") as f:
        for chunk in iter(lambda: f.read(4096), b""):
            hash_md5.update(chunk)
    return hash_md5.hexdigest()


def image_fingerprint(I):
       """similarity measure
       perceptual hash
       image fingerprint
       I is a PIL.Image
       """
       import imagehash
       return str(imagehash.average_hash(I, hash_size=32))


def store_file(db, fn, checksum):
    cur = db.cursor()
    cur.execute('insert into files(filename, hash) values(?,?)', (fn, checksum))
    db.commit()
    return cur.lastrowid


def img_to_png_bytes(I):
    import io
    imgByteArr = io.BytesIO()
    I.save(imgByteArr, format='PNG')
    return imgByteArr.getvalue()


def store_slide(db, fileid, slide):
    cur = db.cursor()
    try:
        cur.execute('insert into slides(fileid, islide, fingerprint, textonly) values(?,?,?,?)',
                (
                    fileid,
                    slide['islide'],
                    slide['fingerprint'],
                    slide['textonly']
                )
        )
    except Exception as err:
        print(err)
        import pdb; pdb.set_trace()

    slideid = cur.lastrowid

    slidebasefn = '%s_%s' % (fileid, slide['islide'])

    # write text
    with open(slidebasefn+'.txt', 'w', encoding="utf-8") as f:
        f.write("\n\n".join(slide['text']))
    if len(slide['text'])>0:
        cur.execute('update slides set text=? where rowid=?', (str(slide['text']), slideid))

    # store thumbenails etc if not just text
    if not slide['textonly'] or True: # FIXME

        # write thumb
        cur.execute('update slides set thumb=? where rowid=?',
                    (
                        #slide['thumb'].tobytes(),
                        img_to_png_bytes(slide['thumb']),
                        slideid
                    )
        )
        #slide['thumb'].save(slidebasefn+'_thumb.png')

        # write hires
        slide['hires'].save(slidebasefn+'.png')
        #cur.execute('update slides set hires=? where rowid=?',
        #            (
        #                img_to_png_bytes(slide['hires']),
        #                slideid
        #            )
        #)

    db.commit()

    return slideid, slidebasefn


def fetch_known_checksums(db):
    cur = db.cursor()
    cur.execute('select hash from files')
    return [row['hash'] for row in cur]


def process1(db, fn, known_checksums):
    logging.info('processing %s...', fn)

    # check that this ppt hasn't been indexed previously based on file md5
    checksum = file_checksum(fn)
    if checksum in known_checksums:
        logging.warning('skipping duplicate %s', fn)
        return

    # write file to library & update known_checksums
    fileid = store_file(db, os.path.abspath(fn), checksum)
    known_checksums.append(checksum)

    # extract: render (thumbnail, hires), text fragments, serialize?
    slides = extract_slides(fn)

    # write slides to library
    for slide in slides:
        slide['fingerprint'] = image_fingerprint(slide['thumb'])
        slideid, outbasefn = store_slide(db, fileid, slide)
        logging.info('wrote slide %s::%s -> %s', fn, slide['islide'], outbasefn)


def main():
    # open existing database if exists else create new
    db = db_connect()

    # get list of folders/files to index
    fns = get_files_to_index()

    # init checksums
    known_checksums = fetch_known_checksums(db)
    logging.info('library knows of %s ppt/pptx files', len(known_checksums))

    # for each file to index
    for fn in fns:
        try:
            process1(db, fn, known_checksums)
        except Exception as err:
            logging.warning('%s failed: %s', fn, err)

if __name__=="__main__":
    main()
