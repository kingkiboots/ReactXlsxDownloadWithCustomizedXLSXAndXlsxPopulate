import TableLayout from 'common/layout/TableLayout';
import React, { useEffect, useState } from 'react';
import { no } from 'util/ComponentUtil';
import { json2XLSX } from 'util/ExcelUtil';
import { numberFormatWithSuffixPeople, numberFormatWithSuffixYear } from 'util/FormatUtil';

const movieList = [
  {
    movieNm: '가디언즈 오브 갤러시 vol.3',
    releasedYear: '2023',
    directorNm: '제임스 건',
    actor: '크리스 프렛',
    attendence: '1000000',
    '3DYn': 'Y',
    text: 'text'
  },
  {
    movieNm: '가디언즈 오브 갤러시 vol.2',
    releasedYear: '2017',
    directorNm: '제임스 건',
    actor: '크리스 프렛',
    attendence: '2000000',
    '3DYn': 'Y',
    text: 'text'
  },
  {
    movieNm: '가디언즈 오브 갤러시 vol.1',
    releasedYear: '2014',
    directorNm: '제임스 건',
    actor: '크리스 프렛',
    attendence: '1500000',
    '3DYn': 'N',
    text: 'text'
  },
  {
    movieNm: '닥터스트레인지',
    releasedYear: '2016',
    directorNm: '스콧 데릭슨',
    actor: '베네딕트 컴버베치',
    attendence: '700000',
    '3DYn': 'N',
    text: 'text'
  },
  {
    movieNm: '닥터스트레인지2',
    releasedYear: '2022',
    directorNm: '샘 레이미',
    actor: '베네딕트 컴버베치',
    attendence: '1200000',
    '3DYn': 'Y',
    text: 'text'
  },
  {
    movieNm: '어벤저스: 엔드게임',
    releasedYear: '2019',
    directorNm: '루소 형제',
    actor: '로버트 다우니 주니어',
    attendence: '10000000',
    '3DYn': 'Y',
    text: 'text'
  },
  {
    movieNm: '어벤저스: 인피니티 워',
    releasedYear: '2018',
    directorNm: '루소 형제',
    actor: '로버트 다우니 주니어',
    attendence: '8000000',
    '3DYn': 'Y',
    text: 'text'
  },
  {
    movieNm: '어벤저스',
    releasedYear: '2012',
    directorNm: '조스 웨던',
    actor: '로다주',
    attendence: '5000000',
    '3DYn': 'N',
    text: 'text'
  },
  {
    movieNm: '토르: 러브 앤 썬더',
    releasedYear: '2022',
    directorNm: '타이카 와이티티',
    actor: '크리스 햄스워드',
    attendence: '1000000',
    '3DYn': 'Y',
    text: 'text'
  },
  {
    movieNm: '로트: 라그나로크',
    releasedYear: '2017',
    directorNm: '타이카 와이티티',
    actor: '크리스 햄스워드',
    attendence: '2000000',
    '3DYn': 'N',
    text: 'text'
  }
];

const layoutHeaderName = '개봉영화 순위';
const SampleTableLayout = () => {
  const [list, setList] = useState([]);

  useEffect(() => {
    fetch();
  }, []);

  const fetch = (isExcel) => {
    const res = movieList;
    const _list = res.map((e) => {
      return {
        movieNm: e.movieNm,
        releasedYear: e.releasedYear,
        directorNm: e.directorNm,
        actor: e.actor,
        attendence: e.attendence,
        '3DYn': e['3DYn'],
        text: e.text
      };
    });
    if (isExcel) json2XLSX(rowDef, _list, `${layoutHeaderName} 엑셀 다운로드`);
    else setList(_list);
  };

  const rowDef = [
    {
      type: no,
      rowSpan: 2,
      labelName: 'No.'
    },
    {
      labelName: '영화이름',
      rowSpan: 2,
      name: 'movieNm'
    },
    {
      labelName: '개봉연도',
      name: 'releasedYear',
      rowSpan: 2,
      format: numberFormatWithSuffixYear
    },
    {
      labelName: '출연진',
      colSpan: 2,
      children: [
        {
          labelName: '영화감독',
          name: 'directorNm'
        },
        {
          labelName: '주연',
          name: 'actor'
        }
      ]
    },
    {
      labelName: '관객수',
      name: 'attendence',
      rowSpan: 2,
      format: numberFormatWithSuffixPeople
    },
    {
      labelName: '3D 여부',
      rowSpan: 2,
      name: '3DYn'
    }
  ];

  const buttonsDef = [
    {
      labelName: '엑셀 다운로드',
      onClick: () => {
        fetch(true);
      }
    }
  ];

  return <TableLayout layoutHeaderName={layoutHeaderName} rowDef={rowDef} list={list} buttonsDef={buttonsDef} />;
};

export default SampleTableLayout;
