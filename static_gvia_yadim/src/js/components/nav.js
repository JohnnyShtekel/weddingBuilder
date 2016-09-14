import React from 'react';
class Nav extends React.Component {


    render() {
        return (
            <nav>
                <div className="nav-wrapper blue-grey darken-1">
                    <a  className="brand-logo right">דוח גבייה</a>
                    <ul id="nav-mobile" className="left hide-on-med-and-down">
                    </ul>
                </div>
            </nav>
        );
    }
}


export default Nav