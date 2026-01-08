from yt_metadata import fetchYoutubeMetadata


def main() -> None:
    url = "https://www.youtube.com/watch?v=TgzP6yhd2vw"
    metadata = fetchYoutubeMetadata(url)
    print(metadata)


if __name__ == "__main__":
    main()
