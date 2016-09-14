import React from 'react';
import { PropTypes } from 'react';


class WelcomePage extends React.Component {
    constructor(props) {
        super(props);


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
                                <li><a href="#" className="round green">עדכוני גבייה<span className="round">עדכוני גבייה אוטומטים בעת הכנסת דו"ח יומי</span></a>
                                </li>
                                <li><a href="#" className="round red">דו"ח מחלקתי<span className="round">הפקת דו"ח מחלקתי שמגיע גם למייל וגם לדפדפן</span></a>
                                </li>
                            </ul>
                        </div>

                    </div>
                </div>
            </div>
        );
    }
}


export default WelcomePage