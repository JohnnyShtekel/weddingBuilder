import React from 'react';
import { RunDepartmentReport } from '../utils/apis';

class DailyDepartmentReport extends React.Component {
    constructor(props) {
        super(props);


    }

    handleSubmit(e) {
        e.preventDefault();
        RunDepartmentReport(document.getElementById("years").value,document.getElementById("mounths").value)
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
                        <span>בחר חודש</span>
                        <div class="input-field col s12">
                            <select id="mounths">
                                <option value="1">ינואר</option>
                                <option value="2">פברואר</option>
                                <option value="3">מרץ</option>
                                <option value="4">אפריל</option>
                                <option value="5">מאי</option>
                                <option value="6">יוני</option>
                                <option value="7">יולי</option>
                                <option value="8">אוגוסט</option>
                                <option value="9">ספטמבר</option>
                                <option value="10">אוקטובר</option>
                                <option value="11">נובמבר</option>
                                <option value="12">דצמבר</option>
                            </select>
                        </div>
                         <span>בחר שנה</span>
                        <div class="input-field col s12">
                            <select id="years" >
                                <option value="2000">2000</option>
                                <option value="2001">2001</option>
                                <option value="2002">2002</option>
                                <option value="2003">2003</option>
                                <option value="2005">2004</option>
                                <option value="2006">2006</option>
                                <option value="2007">2007</option>
                                <option value="2008">2008</option>
                                <option value="2009">2009</option>
                                <option value="2010">2010</option>
                                <option value="2011">2011</option>
                                <option value="2012">2012</option>
                                <option value="2013">2013</option>
                                <option value="2015">2015</option>
                                <option value="2016">2016</option>
                                <option value="2017">2017</option>

                            </select>
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