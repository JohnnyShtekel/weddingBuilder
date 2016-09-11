import React from 'react';
import { RunDepartmentReport } from '../utils/apis';

class DailyDepartmentReport extends React.Component {
    constructor(props) {
        super(props);


    }

    handleSubmit(e) {
        e.preventDefault();
        let currentDate = new Date(new Date(document.getElementById("datepicker").value));
        let day = currentDate.getDay();
        let month = currentDate.getMonth();
        let year = currentDate.getFullYear();
        RunDepartmentReport(day, month, year)
    }

    componentDidMount() {

    }


    componentDidUpdate() {

    }

    componentWillMount(){

    }


    render() {
        return (
            <div className="container cardContainer animated">
                <div className="card blue-grey darken-1">
                    <div className="card-content white-text" dir="rtl">
                        <span className="card-title">ברוכים הבאים</span>
                        <div className="orange-text text-lighten-4">
                            <p>ראשית, בחר תאריך לפי שנה וחודש</p>
                            <p>לאחר מכן, לחץ על כפתור הרץ דו"ח</p>
                        </div>

                         <span>בחר תאריך</span>
                        <div class="input-field col s12">
                           <input type="date" class="datepicker" id="datepicker"/>
                        </div>
                        <div className="container fileInputContainer">
                            <div className="file-field input-field">
                                <div onClick={this.handleSubmit} className="btn white black-text container">
                                    הרץ דוח
                                </div>
                            </div>
                        </div>
                    </div>
                    <div className="card-action" style={{display: 'none'}}>
                        <div className="progress">
                            <div className="indeterminate"></div>
                        </div>
                    </div>
                </div>
            </div>
        );
    }
}


export default DailyDepartmentReport