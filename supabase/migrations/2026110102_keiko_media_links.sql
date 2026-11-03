begin;

update public.home_resources
set
  title = '古今樂明座 Official YouTubeチャンネル',
  description = '公演映像や楽曲から、音・身体の使い方・舞台表現を学べます。',
  destination_url = 'https://m.youtube.com/@Kokongakumeiza',
  updated_at = now()
where source_key = 'kokongakumeiza-youtube';

update public.home_resources
set
  description = '古今樂明座の考え方や稽古の知見を、YouTubeの再生リストで聴けます。',
  image_url = 'https://i.ytimg.com/vi/awLvvl1AvhU/hqdefault.jpg',
  destination_url = 'https://m.youtube.com/watch?v=awLvvl1AvhU&list=PLmk028LkL7RF9l2NIJBfL07O0b1-pBrRJ',
  updated_at = now()
where source_key = 'gakumeiza-radio';

commit;
