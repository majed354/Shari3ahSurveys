window.GRADUATE_SURVEYS_DATA = {
  "title": "استطلاعات الخريجين الخارجية",
  "description": "أُجريت هذه الاستطلاعات عبر شيت مستقل لجمع استجابات الخريجين، ثم أُدرجت نتائجها هنا. تعبّر النتائج عن الخريجين المشاركين في الاستطلاع.",
  "axisNote": "جُمعت الأسئلة تحت محاور لتسهيل العرض؛ الشيت الأصلي لا يتضمن عناوين محاور.",
  "questionnaires": [
    {
      "id": "bachelor",
      "label": "البكالوريوس",
      "sourceFile": "نتائج استطلاعات البكالوريوس.xlsx",
      "questions": [
        {
          "id": "q1",
          "number": 1,
          "label": "بعد تخرجك؛ ما تقييمك العام لجودة التعلم في البرنامج الأكاديمي",
          "axis": "جودة التعلم في البرنامج",
          "kind": "rating",
          "scale": 5
        },
        {
          "id": "q2",
          "number": 2,
          "label": "وضعك الحالي بعد التخرج هو:",
          "axis": "الوضع بعد التخرج",
          "kind": "category",
          "scale": null
        },
        {
          "id": "q3",
          "number": 3,
          "label": "ما درجتك في الاختبارات الوطنية/مهنية بعد التخرج؟ (إذا كنت قد أجريت اختبارًا)",
          "axis": "الاختبارات الوطنية والمهنية",
          "kind": "range",
          "scale": 100
        },
        {
          "id": "q4",
          "number": 4,
          "label": "(للموظفين فقط): ما هو مجال عملك الحالي؟",
          "axis": "الوضع بعد التخرج",
          "kind": "category",
          "scale": null
        },
        {
          "id": "q5",
          "number": 5,
          "label": "(للموظفين فقط): ما درجة تقييم رئيسك في العمل لك، أو كم تقيّم نفسك في عملك الحالي (التقييم من ٥)",
          "axis": "الأداء الوظيفي",
          "kind": "range",
          "scale": 5
        },
        {
          "id": "q6",
          "number": 6,
          "label": "هل لديك أي اقتراحات أو تعليقات",
          "axis": "المقترحات والتعليقات",
          "kind": "text",
          "scale": null
        }
      ]
    },
    {
      "id": "postgraduate",
      "label": "الدراسات العليا",
      "sourceFile": "نتائج استطلاعات الدراسات العليا.xlsx",
      "questions": [
        {
          "id": "q1",
          "number": 1,
          "label": "بعد تخرجك؛ ما تقييمك العام لجودة التعلم في البرنامج الأكاديمي",
          "axis": "جودة التعلم في البرنامج",
          "kind": "rating",
          "scale": 5
        },
        {
          "id": "q2",
          "number": 2,
          "label": "ما تقييمك العام لجودة الإشراف أثناء الرسالة أو المشروع البحثي",
          "axis": "الإشراف العلمي",
          "kind": "rating",
          "scale": 5
        },
        {
          "id": "q3",
          "number": 3,
          "label": "ما مدى رضاك عن مستوى الخدمات المقدمة عموما",
          "axis": "الخدمات المقدمة",
          "kind": "rating",
          "scale": 5
        },
        {
          "id": "q4",
          "number": 4,
          "label": "وضعك الحالي بعد التخرج هو:",
          "axis": "الوضع بعد التخرج",
          "kind": "category",
          "scale": null
        },
        {
          "id": "q5",
          "number": 5,
          "label": "(للموظفين فقط): ما هو مجال عملك الحالي؟",
          "axis": "الوضع بعد التخرج",
          "kind": "category",
          "scale": null
        },
        {
          "id": "q6",
          "number": 6,
          "label": "(للموظفين فقط): ما درجة تقييم رئيسك في العمل لك، أو كم تقيّم نفسك في عملك الحالي (التقييم من ٥)",
          "axis": "الأداء الوظيفي",
          "kind": "range",
          "scale": 5
        },
        {
          "id": "q7",
          "number": 7,
          "label": "هل لديك أي اقتراحات أو تعليقات",
          "axis": "المقترحات والتعليقات",
          "kind": "text",
          "scale": null
        }
      ]
    }
  ],
  "groups": [
    {
      "questionnaireId": "bachelor",
      "programId": "p01",
      "year": "1445",
      "respondents": 7,
      "items": [
        {
          "questionId": "q1",
          "responseCount": 7,
          "scoreTotal": 22.0,
          "categories": [
            {
              "label": "3.0",
              "count": 6
            },
            {
              "label": "4.0",
              "count": 1
            }
          ]
        },
        {
          "questionId": "q2",
          "responseCount": 7,
          "categories": [
            {
              "label": "أبحث عن عمل",
              "count": 4
            },
            {
              "label": "موظف حالياً",
              "count": 3
            }
          ]
        },
        {
          "questionId": "q3",
          "responseCount": 6,
          "scoreTotal": 487.0,
          "categories": [
            {
              "label": "٥٥ - ٥٩",
              "count": 1
            },
            {
              "label": "٧٥ - ٧٩",
              "count": 1
            },
            {
              "label": "٨٠ - ٨٤",
              "count": 1
            },
            {
              "label": "٨٥ - ٨٩",
              "count": 1
            },
            {
              "label": "٩٠ - ٩٤",
              "count": 2
            }
          ]
        },
        {
          "questionId": "q4",
          "responseCount": 3,
          "categories": [
            {
              "label": "استشارات وأبحاث قانونية",
              "count": 1
            },
            {
              "label": "خدمة عملاء",
              "count": 1
            },
            {
              "label": "محامي",
              "count": 1
            }
          ]
        },
        {
          "questionId": "q5",
          "responseCount": 3,
          "scoreTotal": 14.25,
          "categories": [
            {
              "label": "٤.٥ - ٥.٠",
              "count": 3
            }
          ]
        },
        {
          "questionId": "q6",
          "responseCount": 4,
          "categories": []
        }
      ]
    },
    {
      "questionnaireId": "bachelor",
      "programId": "p02",
      "year": "1445",
      "respondents": 7,
      "items": [
        {
          "questionId": "q1",
          "responseCount": 7,
          "scoreTotal": 33.0,
          "categories": [
            {
              "label": "4.0",
              "count": 2
            },
            {
              "label": "5.0",
              "count": 5
            }
          ]
        },
        {
          "questionId": "q2",
          "responseCount": 7,
          "categories": [
            {
              "label": "أبحث عن عمل",
              "count": 5
            },
            {
              "label": "أكمل دراسات عليا",
              "count": 1
            },
            {
              "label": "موظف حالياً",
              "count": 1
            }
          ]
        },
        {
          "questionId": "q3",
          "responseCount": 4,
          "scoreTotal": 253.0,
          "categories": [
            {
              "label": "٦٠ - ٦٤",
              "count": 3
            },
            {
              "label": "٦٥ - ٦٩",
              "count": 1
            }
          ]
        },
        {
          "questionId": "q4",
          "responseCount": 0,
          "categories": []
        },
        {
          "questionId": "q5",
          "responseCount": 1,
          "scoreTotal": 4.245,
          "categories": [
            {
              "label": "٤.٠ - ٤.٤٩",
              "count": 1
            }
          ]
        },
        {
          "questionId": "q6",
          "responseCount": 1,
          "categories": []
        }
      ]
    },
    {
      "questionnaireId": "bachelor",
      "programId": "p02",
      "year": "1446",
      "respondents": 2,
      "items": [
        {
          "questionId": "q1",
          "responseCount": 2,
          "scoreTotal": 7.0,
          "categories": [
            {
              "label": "3.0",
              "count": 1
            },
            {
              "label": "4.0",
              "count": 1
            }
          ]
        },
        {
          "questionId": "q2",
          "responseCount": 2,
          "categories": [
            {
              "label": "أبحث عن عمل",
              "count": 2
            }
          ]
        },
        {
          "questionId": "q3",
          "responseCount": 2,
          "scoreTotal": 134.0,
          "categories": [
            {
              "label": "٥٥ - ٥٩",
              "count": 1
            },
            {
              "label": "٧٥ - ٧٩",
              "count": 1
            }
          ]
        },
        {
          "questionId": "q4",
          "responseCount": 0,
          "categories": []
        },
        {
          "questionId": "q5",
          "responseCount": 0,
          "scoreTotal": 0,
          "categories": []
        },
        {
          "questionId": "q6",
          "responseCount": 2,
          "categories": []
        }
      ]
    },
    {
      "questionnaireId": "bachelor",
      "programId": "p03",
      "year": "1445",
      "respondents": 10,
      "items": [
        {
          "questionId": "q1",
          "responseCount": 10,
          "scoreTotal": 41.0,
          "categories": [
            {
              "label": "2.0",
              "count": 1
            },
            {
              "label": "3.0",
              "count": 3
            },
            {
              "label": "5.0",
              "count": 6
            }
          ]
        },
        {
          "questionId": "q2",
          "responseCount": 10,
          "categories": [
            {
              "label": "أبحث عن عمل",
              "count": 7
            },
            {
              "label": "أكمل دراسات عليا",
              "count": 2
            },
            {
              "label": "موظف حالياً",
              "count": 1
            }
          ]
        },
        {
          "questionId": "q3",
          "responseCount": 8,
          "scoreTotal": 631.0,
          "categories": [
            {
              "label": "٧٠ - ٧٤",
              "count": 4
            },
            {
              "label": "٧٥ - ٧٩",
              "count": 1
            },
            {
              "label": "٨٠ - ٨٤",
              "count": 1
            },
            {
              "label": "٩٠ - ٩٤",
              "count": 2
            }
          ]
        },
        {
          "questionId": "q4",
          "responseCount": 1,
          "categories": [
            {
              "label": "وزارة العدل",
              "count": 1
            }
          ]
        },
        {
          "questionId": "q5",
          "responseCount": 2,
          "scoreTotal": 9.5,
          "categories": [
            {
              "label": "٤.٥ - ٥.٠",
              "count": 2
            }
          ]
        },
        {
          "questionId": "q6",
          "responseCount": 5,
          "categories": []
        }
      ]
    },
    {
      "questionnaireId": "bachelor",
      "programId": "p04",
      "year": "1446",
      "respondents": 13,
      "items": [
        {
          "questionId": "q1",
          "responseCount": 13,
          "scoreTotal": 55.0,
          "categories": [
            {
              "label": "3.0",
              "count": 4
            },
            {
              "label": "4.0",
              "count": 2
            },
            {
              "label": "5.0",
              "count": 7
            }
          ]
        },
        {
          "questionId": "q2",
          "responseCount": 13,
          "categories": [
            {
              "label": "أبحث عن عمل",
              "count": 12
            },
            {
              "label": "أكمل دراسات عليا",
              "count": 1
            }
          ]
        },
        {
          "questionId": "q3",
          "responseCount": 3,
          "scoreTotal": 206.0,
          "categories": [
            {
              "label": "٥٠ - ٥٤",
              "count": 1
            },
            {
              "label": "٦٠ - ٦٤",
              "count": 1
            },
            {
              "label": "٩٠ - ٩٤",
              "count": 1
            }
          ]
        },
        {
          "questionId": "q4",
          "responseCount": 1,
          "categories": [
            {
              "label": "لا",
              "count": 1
            }
          ]
        },
        {
          "questionId": "q5",
          "responseCount": 0,
          "scoreTotal": 0,
          "categories": []
        },
        {
          "questionId": "q6",
          "responseCount": 2,
          "categories": []
        }
      ]
    },
    {
      "questionnaireId": "bachelor",
      "programId": "p04",
      "year": "1445",
      "respondents": 5,
      "items": [
        {
          "questionId": "q1",
          "responseCount": 5,
          "scoreTotal": 17.0,
          "categories": [
            {
              "label": "1.0",
              "count": 1
            },
            {
              "label": "3.0",
              "count": 1
            },
            {
              "label": "4.0",
              "count": 2
            },
            {
              "label": "5.0",
              "count": 1
            }
          ]
        },
        {
          "questionId": "q2",
          "responseCount": 5,
          "categories": [
            {
              "label": "أبحث عن عمل",
              "count": 4
            },
            {
              "label": "موظف حالياً",
              "count": 1
            }
          ]
        },
        {
          "questionId": "q3",
          "responseCount": 3,
          "scoreTotal": 206.0,
          "categories": [
            {
              "label": "٥٠ - ٥٤",
              "count": 1
            },
            {
              "label": "٧٥ - ٧٩",
              "count": 2
            }
          ]
        },
        {
          "questionId": "q4",
          "responseCount": 0,
          "categories": []
        },
        {
          "questionId": "q5",
          "responseCount": 0,
          "scoreTotal": 0,
          "categories": []
        },
        {
          "questionId": "q6",
          "responseCount": 3,
          "categories": []
        }
      ]
    },
    {
      "questionnaireId": "bachelor",
      "programId": "p05",
      "year": "1445",
      "respondents": 6,
      "items": [
        {
          "questionId": "q1",
          "responseCount": 6,
          "scoreTotal": 29.0,
          "categories": [
            {
              "label": "4.0",
              "count": 1
            },
            {
              "label": "5.0",
              "count": 5
            }
          ]
        },
        {
          "questionId": "q2",
          "responseCount": 6,
          "categories": [
            {
              "label": "أبحث عن عمل",
              "count": 4
            },
            {
              "label": "أكمل دراسات عليا",
              "count": 2
            }
          ]
        },
        {
          "questionId": "q3",
          "responseCount": 2,
          "scoreTotal": 139.0,
          "categories": [
            {
              "label": "٥٥ - ٥٩",
              "count": 1
            },
            {
              "label": "٨٠ - ٨٤",
              "count": 1
            }
          ]
        },
        {
          "questionId": "q4",
          "responseCount": 1,
          "categories": [
            {
              "label": "لا",
              "count": 1
            }
          ]
        },
        {
          "questionId": "q5",
          "responseCount": 1,
          "scoreTotal": 4.75,
          "categories": [
            {
              "label": "٤.٥ - ٥.٠",
              "count": 1
            }
          ]
        },
        {
          "questionId": "q6",
          "responseCount": 2,
          "categories": []
        }
      ]
    },
    {
      "questionnaireId": "bachelor",
      "programId": "p01",
      "year": "1446",
      "respondents": 6,
      "items": [
        {
          "questionId": "q1",
          "responseCount": 6,
          "scoreTotal": 22.0,
          "categories": [
            {
              "label": "3.0",
              "count": 4
            },
            {
              "label": "5.0",
              "count": 2
            }
          ]
        },
        {
          "questionId": "q2",
          "responseCount": 6,
          "categories": [
            {
              "label": "أبحث عن عمل",
              "count": 6
            }
          ]
        },
        {
          "questionId": "q3",
          "responseCount": 3,
          "scoreTotal": 251.5,
          "categories": [
            {
              "label": "٦٥ - ٦٩",
              "count": 1
            },
            {
              "label": "٨٥ - ٨٩",
              "count": 1
            },
            {
              "label": "٩٥ - ١٠٠",
              "count": 1
            }
          ]
        },
        {
          "questionId": "q4",
          "responseCount": 1,
          "categories": [
            {
              "label": ".",
              "count": 1
            }
          ]
        },
        {
          "questionId": "q5",
          "responseCount": 0,
          "scoreTotal": 0,
          "categories": []
        },
        {
          "questionId": "q6",
          "responseCount": 4,
          "categories": []
        }
      ]
    },
    {
      "questionnaireId": "bachelor",
      "programId": "p03",
      "year": "1446",
      "respondents": 5,
      "items": [
        {
          "questionId": "q1",
          "responseCount": 5,
          "scoreTotal": 18.0,
          "categories": [
            {
              "label": "2.0",
              "count": 1
            },
            {
              "label": "3.0",
              "count": 1
            },
            {
              "label": "4.0",
              "count": 2
            },
            {
              "label": "5.0",
              "count": 1
            }
          ]
        },
        {
          "questionId": "q2",
          "responseCount": 5,
          "categories": [
            {
              "label": "أبحث عن عمل",
              "count": 2
            },
            {
              "label": "أكمل دراسات عليا",
              "count": 2
            },
            {
              "label": "موظف حالياً",
              "count": 1
            }
          ]
        },
        {
          "questionId": "q3",
          "responseCount": 2,
          "scoreTotal": 169.5,
          "categories": [
            {
              "label": "٧٠ - ٧٤",
              "count": 1
            },
            {
              "label": "٩٥ - ١٠٠",
              "count": 1
            }
          ]
        },
        {
          "questionId": "q4",
          "responseCount": 2,
          "categories": [
            {
              "label": "محامي متدرب",
              "count": 1
            },
            {
              "label": "مستشار قانوني",
              "count": 1
            }
          ]
        },
        {
          "questionId": "q5",
          "responseCount": 2,
          "scoreTotal": 8.49,
          "categories": [
            {
              "label": "٤.٠ - ٤.٤٩",
              "count": 2
            }
          ]
        },
        {
          "questionId": "q6",
          "responseCount": 1,
          "categories": []
        }
      ]
    },
    {
      "questionnaireId": "bachelor",
      "programId": "p05",
      "year": "1446",
      "respondents": 1,
      "items": [
        {
          "questionId": "q1",
          "responseCount": 1,
          "scoreTotal": 5.0,
          "categories": [
            {
              "label": "5.0",
              "count": 1
            }
          ]
        },
        {
          "questionId": "q2",
          "responseCount": 1,
          "categories": [
            {
              "label": "أبحث عن عمل",
              "count": 1
            }
          ]
        },
        {
          "questionId": "q3",
          "responseCount": 0,
          "scoreTotal": 0,
          "categories": []
        },
        {
          "questionId": "q4",
          "responseCount": 0,
          "categories": []
        },
        {
          "questionId": "q5",
          "responseCount": 0,
          "scoreTotal": 0,
          "categories": []
        },
        {
          "questionId": "q6",
          "responseCount": 1,
          "categories": []
        }
      ]
    },
    {
      "questionnaireId": "postgraduate",
      "programId": "p09",
      "year": "1445",
      "respondents": 7,
      "items": [
        {
          "questionId": "q1",
          "responseCount": 7,
          "scoreTotal": 28.0,
          "categories": [
            {
              "label": "3.0",
              "count": 2
            },
            {
              "label": "4.0",
              "count": 3
            },
            {
              "label": "5.0",
              "count": 2
            }
          ]
        },
        {
          "questionId": "q2",
          "responseCount": 7,
          "scoreTotal": 23.0,
          "categories": [
            {
              "label": "1.0",
              "count": 1
            },
            {
              "label": "2.0",
              "count": 2
            },
            {
              "label": "4.0",
              "count": 2
            },
            {
              "label": "5.0",
              "count": 2
            }
          ]
        },
        {
          "questionId": "q3",
          "responseCount": 7,
          "scoreTotal": 20.0,
          "categories": [
            {
              "label": "1.0",
              "count": 1
            },
            {
              "label": "2.0",
              "count": 2
            },
            {
              "label": "3.0",
              "count": 2
            },
            {
              "label": "4.0",
              "count": 1
            },
            {
              "label": "5.0",
              "count": 1
            }
          ]
        },
        {
          "questionId": "q4",
          "responseCount": 7,
          "categories": [
            {
              "label": "أبحث عن عمل",
              "count": 3
            },
            {
              "label": "أكمل دراسات عليا (خاص بالماجستير)",
              "count": 2
            },
            {
              "label": "موظف حالياً",
              "count": 2
            }
          ]
        },
        {
          "questionId": "q5",
          "responseCount": 4,
          "categories": [
            {
              "label": "القضاء",
              "count": 1
            },
            {
              "label": "مؤذن",
              "count": 1
            },
            {
              "label": "محامي",
              "count": 1
            },
            {
              "label": "معلم",
              "count": 1
            }
          ]
        },
        {
          "questionId": "q6",
          "responseCount": 5,
          "scoreTotal": 22.235,
          "categories": [
            {
              "label": "٤.٠ - ٤.٤٩",
              "count": 3
            },
            {
              "label": "٤.٥ - ٥.٠",
              "count": 2
            }
          ]
        },
        {
          "questionId": "q7",
          "responseCount": 2,
          "categories": []
        }
      ]
    },
    {
      "questionnaireId": "postgraduate",
      "programId": "p08",
      "year": "1445",
      "respondents": 2,
      "items": [
        {
          "questionId": "q1",
          "responseCount": 2,
          "scoreTotal": 6.0,
          "categories": [
            {
              "label": "3.0",
              "count": 2
            }
          ]
        },
        {
          "questionId": "q2",
          "responseCount": 2,
          "scoreTotal": 7.0,
          "categories": [
            {
              "label": "2.0",
              "count": 1
            },
            {
              "label": "5.0",
              "count": 1
            }
          ]
        },
        {
          "questionId": "q3",
          "responseCount": 2,
          "scoreTotal": 6.0,
          "categories": [
            {
              "label": "2.0",
              "count": 1
            },
            {
              "label": "4.0",
              "count": 1
            }
          ]
        },
        {
          "questionId": "q4",
          "responseCount": 2,
          "categories": [
            {
              "label": "أبحث عن عمل",
              "count": 1
            },
            {
              "label": "موظف حالياً",
              "count": 1
            }
          ]
        },
        {
          "questionId": "q5",
          "responseCount": 1,
          "categories": [
            {
              "label": "امام مسجد",
              "count": 1
            }
          ]
        },
        {
          "questionId": "q6",
          "responseCount": 1,
          "scoreTotal": 4.245,
          "categories": [
            {
              "label": "٤.٠ - ٤.٤٩",
              "count": 1
            }
          ]
        },
        {
          "questionId": "q7",
          "responseCount": 1,
          "categories": []
        }
      ]
    },
    {
      "questionnaireId": "postgraduate",
      "programId": "p10",
      "year": "1445",
      "respondents": 8,
      "items": [
        {
          "questionId": "q1",
          "responseCount": 8,
          "scoreTotal": 34.0,
          "categories": [
            {
              "label": "2.0",
              "count": 1
            },
            {
              "label": "3.0",
              "count": 1
            },
            {
              "label": "4.0",
              "count": 1
            },
            {
              "label": "5.0",
              "count": 5
            }
          ]
        },
        {
          "questionId": "q2",
          "responseCount": 8,
          "scoreTotal": 31.0,
          "categories": [
            {
              "label": "1.0",
              "count": 1
            },
            {
              "label": "2.0",
              "count": 1
            },
            {
              "label": "3.0",
              "count": 1
            },
            {
              "label": "5.0",
              "count": 5
            }
          ]
        },
        {
          "questionId": "q3",
          "responseCount": 8,
          "scoreTotal": 32.0,
          "categories": [
            {
              "label": "2.0",
              "count": 1
            },
            {
              "label": "3.0",
              "count": 2
            },
            {
              "label": "4.0",
              "count": 1
            },
            {
              "label": "5.0",
              "count": 4
            }
          ]
        },
        {
          "questionId": "q4",
          "responseCount": 8,
          "categories": [
            {
              "label": "أبحث عن عمل",
              "count": 2
            },
            {
              "label": "أكمل دراسات عليا (خاص بالماجستير)",
              "count": 5
            },
            {
              "label": "موظف حالياً",
              "count": 1
            }
          ]
        },
        {
          "questionId": "q5",
          "responseCount": 3,
          "categories": [
            {
              "label": "مساعد إداري",
              "count": 1
            },
            {
              "label": "معلم",
              "count": 1
            },
            {
              "label": "موظف إداري في التعليم",
              "count": 1
            }
          ]
        },
        {
          "questionId": "q6",
          "responseCount": 4,
          "scoreTotal": 19.0,
          "categories": [
            {
              "label": "٤.٥ - ٥.٠",
              "count": 4
            }
          ]
        },
        {
          "questionId": "q7",
          "responseCount": 5,
          "categories": []
        }
      ]
    },
    {
      "questionnaireId": "postgraduate",
      "programId": "p06",
      "year": "1445",
      "respondents": 6,
      "items": [
        {
          "questionId": "q1",
          "responseCount": 6,
          "scoreTotal": 19.0,
          "categories": [
            {
              "label": "2.0",
              "count": 2
            },
            {
              "label": "3.0",
              "count": 2
            },
            {
              "label": "4.0",
              "count": 1
            },
            {
              "label": "5.0",
              "count": 1
            }
          ]
        },
        {
          "questionId": "q2",
          "responseCount": 6,
          "scoreTotal": 16.0,
          "categories": [
            {
              "label": "1.0",
              "count": 1
            },
            {
              "label": "2.0",
              "count": 2
            },
            {
              "label": "3.0",
              "count": 1
            },
            {
              "label": "4.0",
              "count": 2
            }
          ]
        },
        {
          "questionId": "q3",
          "responseCount": 6,
          "scoreTotal": 14.0,
          "categories": [
            {
              "label": "1.0",
              "count": 3
            },
            {
              "label": "2.0",
              "count": 1
            },
            {
              "label": "4.0",
              "count": 1
            },
            {
              "label": "5.0",
              "count": 1
            }
          ]
        },
        {
          "questionId": "q4",
          "responseCount": 6,
          "categories": [
            {
              "label": "أبحث عن عمل",
              "count": 5
            },
            {
              "label": "موظف حالياً",
              "count": 1
            }
          ]
        },
        {
          "questionId": "q5",
          "responseCount": 2,
          "categories": [
            {
              "label": "محاضر متعاون",
              "count": 1
            },
            {
              "label": "محاضر متعاون بجامعة حائل",
              "count": 1
            }
          ]
        },
        {
          "questionId": "q6",
          "responseCount": 2,
          "scoreTotal": 9.5,
          "categories": [
            {
              "label": "٤.٥ - ٥.٠",
              "count": 2
            }
          ]
        },
        {
          "questionId": "q7",
          "responseCount": 6,
          "categories": []
        }
      ]
    },
    {
      "questionnaireId": "postgraduate",
      "programId": "p14",
      "year": "1445",
      "respondents": 5,
      "items": [
        {
          "questionId": "q1",
          "responseCount": 5,
          "scoreTotal": 24.0,
          "categories": [
            {
              "label": "4.0",
              "count": 1
            },
            {
              "label": "5.0",
              "count": 4
            }
          ]
        },
        {
          "questionId": "q2",
          "responseCount": 5,
          "scoreTotal": 21.0,
          "categories": [
            {
              "label": "3.0",
              "count": 1
            },
            {
              "label": "4.0",
              "count": 2
            },
            {
              "label": "5.0",
              "count": 2
            }
          ]
        },
        {
          "questionId": "q3",
          "responseCount": 5,
          "scoreTotal": 20.0,
          "categories": [
            {
              "label": "2.0",
              "count": 1
            },
            {
              "label": "4.0",
              "count": 2
            },
            {
              "label": "5.0",
              "count": 2
            }
          ]
        },
        {
          "questionId": "q4",
          "responseCount": 5,
          "categories": [
            {
              "label": "أبحث عن عمل",
              "count": 3
            },
            {
              "label": "موظف حالياً",
              "count": 2
            }
          ]
        },
        {
          "questionId": "q5",
          "responseCount": 3,
          "categories": [
            {
              "label": "رئيس وحدة البرامج بقسم الوعي الفكري بالتعليم",
              "count": 1
            },
            {
              "label": "معلم بالتعليم وابحث عن العمل بالجامعة.",
              "count": 1
            },
            {
              "label": "وزراة التعليم",
              "count": 1
            }
          ]
        },
        {
          "questionId": "q6",
          "responseCount": 3,
          "scoreTotal": 14.25,
          "categories": [
            {
              "label": "٤.٥ - ٥.٠",
              "count": 3
            }
          ]
        },
        {
          "questionId": "q7",
          "responseCount": 4,
          "categories": []
        }
      ]
    },
    {
      "questionnaireId": "postgraduate",
      "programId": "p11",
      "year": "1446",
      "respondents": 11,
      "items": [
        {
          "questionId": "q1",
          "responseCount": 11,
          "scoreTotal": 49.0,
          "categories": [
            {
              "label": "3.0",
              "count": 1
            },
            {
              "label": "4.0",
              "count": 4
            },
            {
              "label": "5.0",
              "count": 6
            }
          ]
        },
        {
          "questionId": "q2",
          "responseCount": 11,
          "scoreTotal": 38.0,
          "categories": [
            {
              "label": "1.0",
              "count": 3
            },
            {
              "label": "3.0",
              "count": 2
            },
            {
              "label": "4.0",
              "count": 1
            },
            {
              "label": "5.0",
              "count": 5
            }
          ]
        },
        {
          "questionId": "q3",
          "responseCount": 11,
          "scoreTotal": 50.0,
          "categories": [
            {
              "label": "3.0",
              "count": 1
            },
            {
              "label": "4.0",
              "count": 3
            },
            {
              "label": "5.0",
              "count": 7
            }
          ]
        },
        {
          "questionId": "q4",
          "responseCount": 11,
          "categories": [
            {
              "label": "أبحث عن عمل",
              "count": 6
            },
            {
              "label": "أكمل دراسات عليا (خاص بالماجستير)",
              "count": 2
            },
            {
              "label": "موظف حالياً",
              "count": 3
            }
          ]
        },
        {
          "questionId": "q5",
          "responseCount": 3,
          "categories": [
            {
              "label": "متعاقدة عضو هئية تدريس بمعهد الحرم المكي الشريف",
              "count": 1
            },
            {
              "label": "مشرف خدمات بالمسجد النبوي",
              "count": 1
            },
            {
              "label": "معلم",
              "count": 1
            }
          ]
        },
        {
          "questionId": "q6",
          "responseCount": 3,
          "scoreTotal": 13.745,
          "categories": [
            {
              "label": "٤.٠ - ٤.٤٩",
              "count": 1
            },
            {
              "label": "٤.٥ - ٥.٠",
              "count": 2
            }
          ]
        },
        {
          "questionId": "q7",
          "responseCount": 6,
          "categories": []
        }
      ]
    },
    {
      "questionnaireId": "postgraduate",
      "programId": "p08",
      "year": "1446",
      "respondents": 2,
      "items": [
        {
          "questionId": "q1",
          "responseCount": 2,
          "scoreTotal": 6.0,
          "categories": [
            {
              "label": "3.0",
              "count": 2
            }
          ]
        },
        {
          "questionId": "q2",
          "responseCount": 2,
          "scoreTotal": 4.0,
          "categories": [
            {
              "label": "1.0",
              "count": 1
            },
            {
              "label": "3.0",
              "count": 1
            }
          ]
        },
        {
          "questionId": "q3",
          "responseCount": 2,
          "scoreTotal": 5.0,
          "categories": [
            {
              "label": "2.0",
              "count": 1
            },
            {
              "label": "3.0",
              "count": 1
            }
          ]
        },
        {
          "questionId": "q4",
          "responseCount": 2,
          "categories": [
            {
              "label": "أبحث عن عمل",
              "count": 1
            },
            {
              "label": "موظف حالياً",
              "count": 1
            }
          ]
        },
        {
          "questionId": "q5",
          "responseCount": 1,
          "categories": [
            {
              "label": "قانوني",
              "count": 1
            }
          ]
        },
        {
          "questionId": "q6",
          "responseCount": 1,
          "scoreTotal": 4.75,
          "categories": [
            {
              "label": "٤.٥ - ٥.٠",
              "count": 1
            }
          ]
        },
        {
          "questionId": "q7",
          "responseCount": 0,
          "categories": []
        }
      ]
    },
    {
      "questionnaireId": "postgraduate",
      "programId": "p06",
      "year": "1446",
      "respondents": 11,
      "items": [
        {
          "questionId": "q1",
          "responseCount": 11,
          "scoreTotal": 41.0,
          "categories": [
            {
              "label": "2.0",
              "count": 2
            },
            {
              "label": "3.0",
              "count": 3
            },
            {
              "label": "4.0",
              "count": 2
            },
            {
              "label": "5.0",
              "count": 4
            }
          ]
        },
        {
          "questionId": "q2",
          "responseCount": 11,
          "scoreTotal": 44.0,
          "categories": [
            {
              "label": "2.0",
              "count": 1
            },
            {
              "label": "3.0",
              "count": 3
            },
            {
              "label": "4.0",
              "count": 2
            },
            {
              "label": "5.0",
              "count": 5
            }
          ]
        },
        {
          "questionId": "q3",
          "responseCount": 11,
          "scoreTotal": 38.0,
          "categories": [
            {
              "label": "1.0",
              "count": 1
            },
            {
              "label": "2.0",
              "count": 1
            },
            {
              "label": "3.0",
              "count": 4
            },
            {
              "label": "4.0",
              "count": 2
            },
            {
              "label": "5.0",
              "count": 3
            }
          ]
        },
        {
          "questionId": "q4",
          "responseCount": 11,
          "categories": [
            {
              "label": "أبحث عن عمل",
              "count": 6
            },
            {
              "label": "موظف حالياً",
              "count": 5
            }
          ]
        },
        {
          "questionId": "q5",
          "responseCount": 4,
          "categories": [
            {
              "label": "المحاماة",
              "count": 1
            },
            {
              "label": "رقابة",
              "count": 1
            },
            {
              "label": "عسكري",
              "count": 1
            },
            {
              "label": "محاضر بنظام التعاون",
              "count": 1
            }
          ]
        },
        {
          "questionId": "q6",
          "responseCount": 5,
          "scoreTotal": 23.245,
          "categories": [
            {
              "label": "٤.٠ - ٤.٤٩",
              "count": 1
            },
            {
              "label": "٤.٥ - ٥.٠",
              "count": 4
            }
          ]
        },
        {
          "questionId": "q7",
          "responseCount": 6,
          "categories": []
        }
      ]
    },
    {
      "questionnaireId": "postgraduate",
      "programId": "p12",
      "year": "1446",
      "respondents": 1,
      "items": [
        {
          "questionId": "q1",
          "responseCount": 1,
          "scoreTotal": 3.0,
          "categories": [
            {
              "label": "3.0",
              "count": 1
            }
          ]
        },
        {
          "questionId": "q2",
          "responseCount": 1,
          "scoreTotal": 4.0,
          "categories": [
            {
              "label": "4.0",
              "count": 1
            }
          ]
        },
        {
          "questionId": "q3",
          "responseCount": 1,
          "scoreTotal": 2.0,
          "categories": [
            {
              "label": "2.0",
              "count": 1
            }
          ]
        },
        {
          "questionId": "q4",
          "responseCount": 1,
          "categories": [
            {
              "label": "موظف حالياً",
              "count": 1
            }
          ]
        },
        {
          "questionId": "q5",
          "responseCount": 1,
          "categories": [
            {
              "label": "موظف حكومي",
              "count": 1
            }
          ]
        },
        {
          "questionId": "q6",
          "responseCount": 1,
          "scoreTotal": 4.245,
          "categories": [
            {
              "label": "٤.٠ - ٤.٤٩",
              "count": 1
            }
          ]
        },
        {
          "questionId": "q7",
          "responseCount": 0,
          "categories": []
        }
      ]
    },
    {
      "questionnaireId": "postgraduate",
      "programId": "p09",
      "year": "1446",
      "respondents": 8,
      "items": [
        {
          "questionId": "q1",
          "responseCount": 8,
          "scoreTotal": 34.0,
          "categories": [
            {
              "label": "2.0",
              "count": 1
            },
            {
              "label": "3.0",
              "count": 1
            },
            {
              "label": "4.0",
              "count": 1
            },
            {
              "label": "5.0",
              "count": 5
            }
          ]
        },
        {
          "questionId": "q2",
          "responseCount": 8,
          "scoreTotal": 34.0,
          "categories": [
            {
              "label": "1.0",
              "count": 1
            },
            {
              "label": "4.0",
              "count": 2
            },
            {
              "label": "5.0",
              "count": 5
            }
          ]
        },
        {
          "questionId": "q3",
          "responseCount": 8,
          "scoreTotal": 33.0,
          "categories": [
            {
              "label": "2.0",
              "count": 1
            },
            {
              "label": "3.0",
              "count": 1
            },
            {
              "label": "4.0",
              "count": 2
            },
            {
              "label": "5.0",
              "count": 4
            }
          ]
        },
        {
          "questionId": "q4",
          "responseCount": 8,
          "categories": [
            {
              "label": "أبحث عن عمل",
              "count": 5
            },
            {
              "label": "أكمل دراسات عليا (خاص بالماجستير)",
              "count": 1
            },
            {
              "label": "موظف حالياً",
              "count": 2
            }
          ]
        },
        {
          "questionId": "q5",
          "responseCount": 2,
          "categories": [
            {
              "label": "معلم",
              "count": 1
            },
            {
              "label": "معلم في مدرسة خاصة.",
              "count": 1
            }
          ]
        },
        {
          "questionId": "q6",
          "responseCount": 2,
          "scoreTotal": 9.5,
          "categories": [
            {
              "label": "٤.٥ - ٥.٠",
              "count": 2
            }
          ]
        },
        {
          "questionId": "q7",
          "responseCount": 4,
          "categories": []
        }
      ]
    },
    {
      "questionnaireId": "postgraduate",
      "programId": "p15",
      "year": "1446",
      "respondents": 2,
      "items": [
        {
          "questionId": "q1",
          "responseCount": 2,
          "scoreTotal": 8.0,
          "categories": [
            {
              "label": "3.0",
              "count": 1
            },
            {
              "label": "5.0",
              "count": 1
            }
          ]
        },
        {
          "questionId": "q2",
          "responseCount": 2,
          "scoreTotal": 10.0,
          "categories": [
            {
              "label": "5.0",
              "count": 2
            }
          ]
        },
        {
          "questionId": "q3",
          "responseCount": 2,
          "scoreTotal": 8.0,
          "categories": [
            {
              "label": "3.0",
              "count": 1
            },
            {
              "label": "5.0",
              "count": 1
            }
          ]
        },
        {
          "questionId": "q4",
          "responseCount": 2,
          "categories": [
            {
              "label": "أبحث عن عمل",
              "count": 1
            },
            {
              "label": "موظف حالياً",
              "count": 1
            }
          ]
        },
        {
          "questionId": "q5",
          "responseCount": 1,
          "categories": [
            {
              "label": "أستاذ مساعد بقسم القراءات في جامعة أم القرى",
              "count": 1
            }
          ]
        },
        {
          "questionId": "q6",
          "responseCount": 1,
          "scoreTotal": 4.75,
          "categories": [
            {
              "label": "٤.٥ - ٥.٠",
              "count": 1
            }
          ]
        },
        {
          "questionId": "q7",
          "responseCount": 2,
          "categories": []
        }
      ]
    },
    {
      "questionnaireId": "postgraduate",
      "programId": "p15",
      "year": "1445",
      "respondents": 1,
      "items": [
        {
          "questionId": "q1",
          "responseCount": 1,
          "scoreTotal": 3.0,
          "categories": [
            {
              "label": "3.0",
              "count": 1
            }
          ]
        },
        {
          "questionId": "q2",
          "responseCount": 1,
          "scoreTotal": 5.0,
          "categories": [
            {
              "label": "5.0",
              "count": 1
            }
          ]
        },
        {
          "questionId": "q3",
          "responseCount": 1,
          "scoreTotal": 1.0,
          "categories": [
            {
              "label": "1.0",
              "count": 1
            }
          ]
        },
        {
          "questionId": "q4",
          "responseCount": 1,
          "categories": [
            {
              "label": "أبحث عن عمل",
              "count": 1
            }
          ]
        },
        {
          "questionId": "q5",
          "responseCount": 1,
          "categories": [
            {
              "label": "معلم في وزارة التعليم",
              "count": 1
            }
          ]
        },
        {
          "questionId": "q6",
          "responseCount": 1,
          "scoreTotal": 4.75,
          "categories": [
            {
              "label": "٤.٥ - ٥.٠",
              "count": 1
            }
          ]
        },
        {
          "questionId": "q7",
          "responseCount": 1,
          "categories": []
        }
      ]
    },
    {
      "questionnaireId": "postgraduate",
      "programId": "p13",
      "year": "1446",
      "respondents": 1,
      "items": [
        {
          "questionId": "q1",
          "responseCount": 1,
          "scoreTotal": 3.0,
          "categories": [
            {
              "label": "3.0",
              "count": 1
            }
          ]
        },
        {
          "questionId": "q2",
          "responseCount": 1,
          "scoreTotal": 2.0,
          "categories": [
            {
              "label": "2.0",
              "count": 1
            }
          ]
        },
        {
          "questionId": "q3",
          "responseCount": 1,
          "scoreTotal": 3.0,
          "categories": [
            {
              "label": "3.0",
              "count": 1
            }
          ]
        },
        {
          "questionId": "q4",
          "responseCount": 1,
          "categories": [
            {
              "label": "أبحث عن عمل",
              "count": 1
            }
          ]
        },
        {
          "questionId": "q5",
          "responseCount": 0,
          "categories": []
        },
        {
          "questionId": "q6",
          "responseCount": 0,
          "scoreTotal": 0,
          "categories": []
        },
        {
          "questionId": "q7",
          "responseCount": 0,
          "categories": []
        }
      ]
    },
    {
      "questionnaireId": "postgraduate",
      "programId": "p10",
      "year": "1446",
      "respondents": 10,
      "items": [
        {
          "questionId": "q1",
          "responseCount": 10,
          "scoreTotal": 49.0,
          "categories": [
            {
              "label": "4.0",
              "count": 1
            },
            {
              "label": "5.0",
              "count": 9
            }
          ]
        },
        {
          "questionId": "q2",
          "responseCount": 10,
          "scoreTotal": 41.0,
          "categories": [
            {
              "label": "2.0",
              "count": 2
            },
            {
              "label": "3.0",
              "count": 1
            },
            {
              "label": "4.0",
              "count": 1
            },
            {
              "label": "5.0",
              "count": 6
            }
          ]
        },
        {
          "questionId": "q3",
          "responseCount": 10,
          "scoreTotal": 44.0,
          "categories": [
            {
              "label": "3.0",
              "count": 3
            },
            {
              "label": "5.0",
              "count": 7
            }
          ]
        },
        {
          "questionId": "q4",
          "responseCount": 10,
          "categories": [
            {
              "label": "أبحث عن عمل",
              "count": 6
            },
            {
              "label": "أكمل دراسات عليا (خاص بالماجستير)",
              "count": 2
            },
            {
              "label": "موظف حالياً",
              "count": 2
            }
          ]
        },
        {
          "questionId": "q5",
          "responseCount": 2,
          "categories": [
            {
              "label": "معلمه في قطاع خاص",
              "count": 1
            },
            {
              "label": "هيئة الزكاة والضريبة والجمارك",
              "count": 1
            }
          ]
        },
        {
          "questionId": "q6",
          "responseCount": 2,
          "scoreTotal": 9.5,
          "categories": [
            {
              "label": "٤.٥ - ٥.٠",
              "count": 2
            }
          ]
        },
        {
          "questionId": "q7",
          "responseCount": 4,
          "categories": []
        }
      ]
    },
    {
      "questionnaireId": "postgraduate",
      "programId": "p14",
      "year": "1446",
      "respondents": 4,
      "items": [
        {
          "questionId": "q1",
          "responseCount": 4,
          "scoreTotal": 17.0,
          "categories": [
            {
              "label": "3.0",
              "count": 1
            },
            {
              "label": "4.0",
              "count": 1
            },
            {
              "label": "5.0",
              "count": 2
            }
          ]
        },
        {
          "questionId": "q2",
          "responseCount": 4,
          "scoreTotal": 18.0,
          "categories": [
            {
              "label": "3.0",
              "count": 1
            },
            {
              "label": "5.0",
              "count": 3
            }
          ]
        },
        {
          "questionId": "q3",
          "responseCount": 4,
          "scoreTotal": 20.0,
          "categories": [
            {
              "label": "5.0",
              "count": 4
            }
          ]
        },
        {
          "questionId": "q4",
          "responseCount": 4,
          "categories": [
            {
              "label": "أبحث عن عمل",
              "count": 1
            },
            {
              "label": "موظف حالياً",
              "count": 3
            }
          ]
        },
        {
          "questionId": "q5",
          "responseCount": 3,
          "categories": [
            {
              "label": "التعليم",
              "count": 1
            },
            {
              "label": "معلم",
              "count": 1
            },
            {
              "label": "معلم متقدم بوزارة التعليم",
              "count": 1
            }
          ]
        },
        {
          "questionId": "q6",
          "responseCount": 3,
          "scoreTotal": 14.25,
          "categories": [
            {
              "label": "٤.٥ - ٥.٠",
              "count": 3
            }
          ]
        },
        {
          "questionId": "q7",
          "responseCount": 3,
          "categories": []
        }
      ]
    }
  ],
  "years": [
    "1446",
    "1445"
  ],
  "programs": [
    {
      "id": "p01",
      "name": "الأنظمة",
      "degree": "البكالوريوس",
      "dept": "الأنظمة"
    },
    {
      "id": "p02",
      "name": "الدراسات الإسلامية",
      "degree": "البكالوريوس",
      "dept": "الدراسات الإسلامية"
    },
    {
      "id": "p03",
      "name": "الشريعة",
      "degree": "البكالوريوس",
      "dept": "الشريعة"
    },
    {
      "id": "p04",
      "name": "القرآن وعلومه",
      "degree": "البكالوريوس",
      "dept": "القراءات"
    },
    {
      "id": "p05",
      "name": "القراءات",
      "degree": "البكالوريوس",
      "dept": "القراءات"
    },
    {
      "id": "p06",
      "name": "القانون",
      "degree": "الماجستير",
      "dept": "الأنظمة"
    },
    {
      "id": "p08",
      "name": "أصول الفقه",
      "degree": "الماجستير",
      "dept": "الشريعة"
    },
    {
      "id": "p09",
      "name": "الفقه",
      "degree": "الماجستير",
      "dept": "الشريعة"
    },
    {
      "id": "p10",
      "name": "الدراسات القرآنية المعاصرة",
      "degree": "الماجستير",
      "dept": "القراءات"
    },
    {
      "id": "p11",
      "name": "القراءات",
      "degree": "الماجستير",
      "dept": "القراءات"
    },
    {
      "id": "p12",
      "name": "أصول الفقه",
      "degree": "الدكتوراه",
      "dept": "الشريعة"
    },
    {
      "id": "p13",
      "name": "الفقه",
      "degree": "الدكتوراه",
      "dept": "الشريعة"
    },
    {
      "id": "p14",
      "name": "الدراسات القرآنية",
      "degree": "الدكتوراه",
      "dept": "القراءات"
    },
    {
      "id": "p15",
      "name": "القراءات",
      "degree": "الدكتوراه",
      "dept": "القراءات"
    }
  ]
};
