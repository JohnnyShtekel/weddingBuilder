import React from 'react';
import { PropTypes } from 'react';


class WelcomePage extends React.Component {
    constructor(props) {
        super(props);


    }


    moveTorunDailyDepartmentReport(){
        this.context.history.pushState(null, '/api/v1/gvia-yadim-report/department-report/');
    }

    moveToUploadPage(){
        this.context.history.pushState(null, '/api/v1/gvia-yadim-report/upload-page/');
    }

    render() {

        return (
            <div className="container cardContainer animated">
                <div className="card blue-grey darken-1">
                    <div className="card-content white-text" dir="rtl">
                        <span className="card-title">ברוכים הבאים</span>
                        <div className="orange-text text-lighten-4">
                            <p>נא בחר את אחת האופציות</p>
                        </div>
                        <div className="row">
                            <ul>
                                <li><a onClick={this.moveToUploadPage.bind(this)} className="round green">עדכוני גבייה<span className="round">עדכוני גבייה אוטומטים בעת הכנסת דו"ח יומי</span></a>
                                </li>
                                <li><a onClick={this.moveTorunDailyDepartmentReport.bind(this)} className="round red">דו"ח מחלקתי<span className="round">הפקת דו"ח מחלקתי שמגיע גם למייל וגם לדפדפן</span></a>
                                </li>
                            </ul>
                        </div>

                    </div>
                </div>
            </div>
        );
    }
}

WelcomePage.contextTypes = {
    history: React.PropTypes.object
};

export default WelcomePage