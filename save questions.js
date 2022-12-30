let fst_year_forms = {
    aqida: "1Gqqtw34JkaODRZHXKGmW_untLYWq8-lhU8_hM62PapA",
}
let fst_year_fst_term_xls = {
    study_year_id: 1,
    term_id: 1,
    subjects: [
        {
            name: "العقيدة",
            sheet_id: "1ym5xf67pp_A4nrR648cI-6Uvy2c_K0V97FrppBSbCBc"
        }
    ]
}

let form = FormApp.openById(fst_year_forms.aqida);
// Open a sheet by ID.
let sheet = SpreadsheetApp.openById(fst_year_fst_term_xls.subjects[0].sheet_id).getSheets()[0];

let formName = form.getTitle()

// variables for putting the questions and answers in the right position
let question_position = 1;
let right_answer_index = 17

// main function to run
function getFormValues() {
    form.getItems().forEach(callback);
}

{
    let answers_position = 1;

    //title
    sheet.getRange(question_position, answers_position++).setValue("title");
    //points
    sheet.getRange(question_position, answers_position++).setValue("points");
    //lesson or subject
    sheet.getRange(question_position, answers_position++).setValue("lesson or subject");
    //path
    sheet.getRange(question_position, answers_position++).setValue("path");
    //term
    sheet.getRange(question_position, answers_position++).setValue("TERM");
    //subject
    sheet.getRange(question_position, answers_position++).setValue("subject");
    //lesson
    sheet.getRange(question_position, answers_position++).setValue("lesson");
    //question type
    sheet.getRange(question_position, answers_position++).setValue("type");
    //description
    sheet.getRange(question_position, answers_position++).setValue("description");
    //choices count
    sheet.getRange(question_position, answers_position++).setValue("choices count");
    //coices
    sheet.getRange(question_position, answers_position++).setValue("choice1");
    sheet.getRange(question_position, answers_position++).setValue("choice2");
    sheet.getRange(question_position, answers_position++).setValue("choice3");
    sheet.getRange(question_position, answers_position++).setValue("choice4");
    sheet.getRange(question_position, answers_position++).setValue("choice5");
    sheet.getRange(question_position, answers_position++).setValue("choice6");
    //right answer
    sheet.getRange(question_position, answers_position++).setValue("answer");
    question_position++;
}

// Iterate over all questions
function callback(el) {
    let answers_position = 1;
    let question;
    let title;
    let choices;
    let type = el.getType()
    let points;
    switch (el.getType()) {
        case FormApp.ItemType.MULTIPLE_CHOICE:
            type = "MULTIPLE_CHOICE"
            question = el.asMultipleChoiceItem();
            points = question.getPoints();
            if (points === 0) {
                return;
            }
            title = question.getTitle();

        {
            if (title === "المذهب" ||
                title === "تأكدت من كتابة البريد الإلكتروني صحيحًا" ||
                title === "تأكدت من كتابة الاسم كاملًا كما هو في الأوراق الرسمية باللغة العربية" ||
                title === "تأكدت من كتابة الرقم الجامعي كما هو في الأوراق الرسمية" ||
                title === "دخولك الاختبار أكثر من مرة يعرضك لخسارة جميع علاماتك" ||
                title === "أتعهد بعدم فتح الكتاب أو أي مصدر آخر أثناء الاختبار، وعدم نشر أو مناقشة أسئلة الاختبار إلا بعد انتهاء الوقت المحدد  والله على ما أقول  شهيد." ||
                title === "النوع"
            ) {
                return;
            }
        }
            choices = question.getChoices();

            //title
            sheet.getRange(question_position, answers_position++).setValue(title);
            //points
            sheet.getRange(question_position, answers_position++).setValue(points);
            //lesson or subject
            sheet.getRange(question_position, answers_position++).setValue("subject");
            //path
            sheet.getRange(question_position, answers_position++).setValue(fst_year_fst_term_xls.study_year_id);
            //term
            sheet.getRange(question_position, answers_position++).setValue(fst_year_fst_term_xls.term_id);
            //subject
            sheet.getRange(question_position, answers_position++).setValue(fst_year_fst_term_xls.subjects[0].name);
            //lesson
            sheet.getRange(question_position, answers_position++).setValue("");


            if (choices.length === 2) {
                if ((choices[0].getValue().includes("صح") && choices[1].getValue().includes("خطأ")) ||
                    (choices[1].getValue().includes("صح") && choices[0].getValue().includes("خطأ"))) {
                    type = "binary"
                    //question type
                    sheet.getRange(question_position, answers_position++).setValue(type);
                    //description
                    sheet.getRange(question_position, answers_position++).setValue("");
                    //choices count
                    sheet.getRange(question_position, answers_position++).setValue("");
                    choices.forEach(function (choice) {
                        if (choice.isCorrectAnswer()) {
                            if (choice.getValue().includes("صح")) {
                                sheet.getRange(question_position, right_answer_index).setValue("صح");
                            } else {
                                sheet.getRange(question_position, right_answer_index).setValue("خطأ");
                            }
                        }
                    });
                    question_position++;
                    return;
                }
            }


            //question type
            sheet.getRange(question_position, answers_position++).setValue(type);
            //description
            sheet.getRange(question_position, answers_position++).setValue("");
            //choices count
            sheet.getRange(question_position, answers_position++).setValue(choices.length);
            choices.forEach(function (choice, index) {
                sheet.getRange(question_position, answers_position++).setValue(choice.getValue());
                if (choice.isCorrectAnswer()) {
                    sheet.getRange(question_position, right_answer_index).setValue(index + 1);
                }
            });

            question_position++;
            break;
        case FormApp.ItemType.CHECKBOX:
            type = "CHECKBOX"
            question = el.asCheckboxItem()
            points = question.getPoints();
            if (points === 0) {
                return;
            }

            title = question.getTitle();
            choices = question.getChoices();


            //title
            sheet.getRange(question_position, answers_position++).setValue(title);
            //points
            sheet.getRange(question_position, answers_position++).setValue(points);
            //lesson or subject
            sheet.getRange(question_position, answers_position++).setValue("subject");
            //path
            sheet.getRange(question_position, answers_position++).setValue(fst_year_fst_term_xls.study_year_id);
            //term
            sheet.getRange(question_position, answers_position++).setValue(fst_year_fst_term_xls.term_id);
            //subject
            sheet.getRange(question_position, answers_position++).setValue(fst_year_fst_term_xls.subjects[0].name);
            //lesson
            sheet.getRange(question_position, answers_position++).setValue("");
            //question type
            sheet.getRange(question_position, answers_position++).setValue(type);
            //description
            sheet.getRange(question_position, answers_position++).setValue("");
            //choices count
            sheet.getRange(question_position, answers_position++).setValue(choices.length);
            choices.forEach(function (choice, index) {
                sheet.getRange(question_position, answers_position++).setValue(choice.getValue());
                if (choice.isCorrectAnswer()) {
                    let correct_answer = sheet.getRange(question_position, right_answer_index).getValue();
                    sheet.getRange(question_position, right_answer_index).setValue(correct_answer + (index + 1) + ",");
                }
            });

            question_position++;
            break;
    }
}




