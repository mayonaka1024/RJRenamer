import asyncio
from dlsite_async import DlsiteAPI
import xml.etree.ElementTree as ET
import sys
import os
import zipfile
import tempfile
import bs4
import requests
import re

def main():
    if len(sys.argv) > 1:
        file_paths = sys.argv[1:]
    else:
        print("ファイルパスを引数で指定してください")
        sys.exit()

    for file_path in file_paths:
        try:
            file_name = os.path.basename(file_path)
            name_without_ext = os.path.splitext(file_name)[0]
            print(file_name)

            code = name_without_ext

            if code.startswith("RJ"):
                comic_info = asyncio.run(fetch_dlsite_metadata(code))
            elif code.startswith("d_"):
                comic_info = asyncio.run(fetch_fanza_metadata(code))
            else:
                print(f"ファイル名が作品コードの形式ではありません: {file_name}")
                continue

            if comic_info is None:
                print(f"情報取得に失敗しました: {file_name}")
                continue

            limited_title = comic_info['Title'][:30] if comic_info['Title'] else ''
            new_file_name = f"({comic_info['FormatName']}) [{comic_info['Circle']} ({comic_info['author']})] {limited_title} [{code}].zip"
            new_file_name = sanitize_filename(new_file_name)
            print(new_file_name)

            # ジャンル（作品形式）が CG またはコミックのときだけ ComicInfo.xml を作成・格納する
            if is_comic_genre(comic_info.get('FormatName', '')):
                write_comicinfo_xml(comic_info, "ComicInfo.xml")
                add_file_to_zip(file_path, "ComicInfo.xml")
                rename_file(file_path, new_file_name)
                os.remove("ComicInfo.xml")
            else:
                # それ以外はファイル名の変更だけ行う
                rename_file(file_path, new_file_name)
        except Exception as e:
            print(f"エラー発生: {file_path} → {e}")
            continue


def is_comic_genre(format_name):
    """作品形式が CG またはコミックの場合に True を返す。

    DLsite: worktype_to_label 変換後の「同人CG集」「コミック」などを想定。
    FANZA: 「CG・イラスト」「コミック」など、work_type の文字列をそのまま想定。
    """
    if not format_name:
        return False
    name = str(format_name)
    # 「CG」または「コミック」を含むかで判定
    return ("CG" in name) or ("コミック" in name)


async def fetch_fanza_metadata(code):
    fanza_url= "https://www.dmm.co.jp/dc/doujin/-/detail/=/cid="
    code_URL= fanza_url + code
    cookie = {'age_check_done': '1'}
    webpage = requests.get(code_URL,cookies=cookie)
    soup = bs4.BeautifulSoup(webpage.content, "html.parser")

    author = ""
    work_type = ""
    publish_date = ""
    genre_names = []
    circle_name = soup.find("a", class_="circleName__txt").text.strip()
    title = soup.find("h1", class_="productTitle__txt").text
    import re
    if title:
        title = title.replace('\n', '').replace('\r', '').strip()
        title = re.sub(r'^【.*?OFF】', '', title).lstrip()

    for item in soup.find_all("div", class_="productInformation__item"):
        dt = item.find("dt", class_="informationList__ttl")
        dd = item.find("dd", class_="informationList__txt")
        if dt and dd and dd.a:
            if dt.text.strip() == "作者":
                author = dd.a.text.strip()
            elif dt.text.strip() == "作品形式":
                work_type = dd.a.text.strip()
            elif dt.text.strip() == "配信開始日":
                publish_date = dd.a.text.strip()

    for a in soup.select("ul.genreTagList a.genreTag__txt"):
        genre_names.append(a.text.strip())


    print("タイトル:", title)
    print("サークル名:", circle_name)
    print("作者:", author)
    print("作品形式:", work_type)
    print("配信開始日:", publish_date)
    print("ジャンル:", genre_names)

    # ジャンル名の頭に "Ganre:" を付与（フォールバック時も付与）
    prefixed_genres = [f"Ganre:{g}" for g in genre_names]
    # サークル名の頭に "group:" を付与
    prefixed_circle = f"group:{circle_name}"
    # 作者名の頭に "artist:" を付与
    prefixed_authors = [f"artist:{a.strip()}" for a in author.split(",") if a.strip()] if author else []
    # 作品コードの頭に "code:" を付与
    prefixed_code = f"code:{code}"
    # 作品形式の頭に "Format:" を付与
    prefixed_format = f"Format:{work_type}" if work_type else None
    # サークル名・作者・ジャンル・作品コード・作品形式を Tags にまとめる
    tag_parts = [prefixed_circle] + prefixed_authors + prefixed_genres + [prefixed_code]
    if prefixed_format:
        tag_parts.append(prefixed_format)
    tags = ", ".join(tag_parts)

    comic_info = {
    "Title": title,
    "author": author,
    "Circle": circle_name,
    "Publisher": "FANZA",
    "Year": publish_date[:4],
    "Month": publish_date[5:7],
    "Day": publish_date[8:10],
    "Language": "Japanese",
    "FormatName": work_type,  # ファイル名生成用（ComicInfo.xmlには出力しない）
    "summary": "",
    "Tags": tags
    }
    return comic_info


def worktype_to_label(worktype):
    mapping = {
        "ACTION": "アクション",
        "ACN": "アクション",
        "ADVENTURE": "アドベンチャー",
        "ADV": "アドベンチャー",
        "QUIZ": "クイズ",
        "QIZ": "クイズ",
        "CG_ILLUSTRATIONS": "同人CG集",
        "ICG": "同人CG集",
        "DIGITAL_NOVEL": "デジタルノベル",
        "DNV": "デジタルノベル",
        "GEKIGA": "劇画",
        "SCM": "劇画",
        "ILLUST_MATERIALS": "イラスト素材",
        "IMT": "イラスト素材",
        "MANGA": "コミック",
        "MNG": "コミック",
        "MISCELLANEOUS": "その他",
        "ET3": "その他",
        "MISCELLANEOUS_GAME": "その他ゲーム",
        "ETC": "その他ゲーム",
        "MUSIC": "音楽",
        "MUS": "音楽",
        "MUSIC_MATERIALS": "音楽素材",
        "AMT": "音楽素材",
        "NOVEL": "ノベル",
        "NRE": "ノベル",
        "PUBLICATION": "出版物",
        "PBC": "出版物",
        "PUZZLE": "パズル",
        "PZL": "パズル",
        "ROLE_PLAYING": "ロールプレイング",
        "RPG": "ロールプレイング",
        "SHOOTING": "シューティング",
        "STG": "シューティング",
        "SIMULATION": "シミュレーション",
        "SLN": "シミュレーション",
        "TABLE": "テーブルゲーム",
        "TBL": "テーブルゲーム",
        "TOOLS_ACCESSORIES": "ツール・アクセサリ",
        "TOL": "ツール・アクセサリ",
        "TYPING": "タイピング",
        "TYP": "タイピング",
        "VIDEO": "動画",
        "MOV": "動画",
        "VOICE_ASMR": "ASMR",
        "SOU": "ASMR",
        "VOICED_COMIC": "ボイスコミック",
        "VCM": "ボイスコミック",
        "WEBTOON": "ウェブトゥーン",
        "WBT": "ウェブトゥーン"
    }
    return mapping.get(str(worktype), str(worktype))

async def fetch_dlsite_metadata(rj_code):
    async with DlsiteAPI() as api:
        result = await api.get_work(rj_code)
        author_list = getattr(result, "author", [])
        if author_list is None:
            author_list = []
        if result:
            circle_name = getattr(result, "circle", "")
            genre_names = getattr(result, "genre", []) or []
            # ジャンル名の頭に "Ganre:" を付与（フォールバック時も付与）
            prefixed_genres = [f"Ganre:{g}" for g in genre_names]
            # サークル名の頭に "group:" を付与
            prefixed_circle = f"group:{circle_name}"
            # 作者名の頭に "artist:" を付与
            prefixed_authors = [f"artist:{a}" for a in author_list if a]
            # 作品コードの頭に "code:" を付与
            prefixed_code = f"code:{rj_code}"
            # 作品形式のラベル
            format_name = worktype_to_label(result.work_type.name) if hasattr(result, "work_type") else ""
            # 作品形式の頭に "Format:" を付与
            prefixed_format = f"Format:{format_name}" if format_name else None
            # サークル名・作者・ジャンル・作品コード・作品形式を Tags にまとめる
            tag_parts = [prefixed_circle] + prefixed_authors + prefixed_genres + [prefixed_code]
            if prefixed_format:
                tag_parts.append(prefixed_format)
            tags = ", ".join(tag_parts)

            comic_info = {
            "Title": getattr(result, "work_name", ""),
            "author": ", ".join(author_list),
            "Circle": circle_name,
            "Publisher": "DLsite",
            "Year": getattr(result.regist_date, "year", "") if hasattr(result, "regist_date") and result.regist_date else "",
            "Month": getattr(result.regist_date, "month", "") if hasattr(result, "regist_date") and result.regist_date else "",
            "Day": getattr(result.regist_date, "day", "") if hasattr(result, "regist_date") and result.regist_date else "",
            "Language": "Japanese",
            "FormatName":  format_name,  # ファイル名生成用（ComicInfo.xmlには出力しない）
            "summary": getattr(result, "description", ""),
            "Tags": tags

        }
            return comic_info

def write_comicinfo_xml(info: dict, output_path: str):
    # ファイル名生成専用の内部キーはComicInfo.xmlには出力しない
    exclude_keys = {"FormatName"}
    root = ET.Element("ComicInfo")
    for key, value in info.items():
        if key in exclude_keys:
            continue
        elem = ET.SubElement(root, key)
        elem.text = str(value) if value is not None else ""
    tree = ET.ElementTree(root)
    tree.write(output_path, encoding="utf-8", xml_declaration=True)

def rename_file(src_path, new_name):
    dir_path = os.path.dirname(src_path)
    dst_path = os.path.join(dir_path, new_name)
    os.rename(src_path, dst_path)
    return dst_path

def add_file_to_zip(zip_path, file_path):
    arcname = os.path.basename(file_path)

    with zipfile.ZipFile(zip_path, 'r') as zipf:
        existing_names = zipf.namelist()

    # 既存に同名エントリが無ければ単純に追加
    if arcname not in existing_names:
        with zipfile.ZipFile(zip_path, 'a', zipfile.ZIP_DEFLATED) as zipf:
            zipf.write(file_path, arcname)
        return

    # 既存の同名エントリがある場合は上書き（既存エントリを除いてZIPを作り直す）
    print(f"{arcname} は既にZIPに存在します。上書きします。")
    dir_path = os.path.dirname(os.path.abspath(zip_path))
    tmp_fd, tmp_path = tempfile.mkstemp(suffix=".zip", dir=dir_path)
    os.close(tmp_fd)
    try:
        with zipfile.ZipFile(zip_path, 'r') as zin, \
             zipfile.ZipFile(tmp_path, 'w', zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                # 上書き対象の同名エントリはコピーしない
                if item.filename == arcname:
                    continue
                zout.writestr(item, zin.read(item.filename))
            # 新しいComicInfo.xmlを追加
            zout.write(file_path, arcname)
        os.replace(tmp_path, zip_path)
    except Exception:
        # 失敗時は一時ファイルを掃除
        if os.path.exists(tmp_path):
            os.remove(tmp_path)
        raise

def sanitize_filename(filename):
    invalid_chars = {
        '"': '”', '<': '＜', '>': '＞', '|': '｜',
        '*': '＊', '?': '？', '\\': '￥', '/': '／', ':': '：'
    }

    for k, v in invalid_chars.items():
        filename = filename.replace(k, v)
    # 連続空白を1つにまとめる
    filename = re.sub(r'\s+', ' ', filename)
    # 先頭・末尾の空白とピリオドを除去
    filename = filename.strip(' .')
    # フルパス長制限（255文字）
    # 呼び出し元でdir_pathと組み合わせて使う想定
    import inspect
    frame = inspect.currentframe().f_back
    dir_path = None
    if 'src_path' in frame.f_locals:
        dir_path = os.path.dirname(frame.f_locals['src_path'])
    elif 'file_path' in frame.f_locals:
        dir_path = os.path.dirname(frame.f_locals['file_path'])
    if dir_path:
        full_path = os.path.join(dir_path, filename)
        max_len = 255
        if len(full_path) > max_len:
            base, ext = os.path.splitext(filename)
            # 拡張子分を残して短縮
            allowed = max_len - len(dir_path) - len(os.sep) - len(ext)
            if allowed < 1:
                allowed = 1
            base = base[:allowed]
            filename = base.strip(' .') + ext
    else:
        # ディレクトリ情報が取れない場合は255文字制限
        if len(filename) > 255:
            base, ext = os.path.splitext(filename)
            base = base[:255-len(ext)]
            filename = base.strip(' .') + ext
    return filename


if __name__ == "__main__":
    main()
