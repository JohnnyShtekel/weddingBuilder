import React from 'react';
import Nav from './components/nav'

class App extends React.Component {
    constructor(props) {
        super(props);
    }


    render() {
        return (
            <div id="app-id">
                <Nav/>
                {this.props.children}
            </div>

        )
    }

}
export default App;