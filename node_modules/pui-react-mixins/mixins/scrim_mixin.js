/*(c) Copyright 2015 Pivotal Software, Inc. All Rights Reserved.*/
'use strict';

var _createClass = function () { function defineProperties(target, props) { for (var i = 0; i < props.length; i++) { var descriptor = props[i]; descriptor.enumerable = descriptor.enumerable || false; descriptor.configurable = true; if ("value" in descriptor) descriptor.writable = true; Object.defineProperty(target, descriptor.key, descriptor); } } return function (Constructor, protoProps, staticProps) { if (protoProps) defineProperties(Constructor.prototype, protoProps); if (staticProps) defineProperties(Constructor, staticProps); return Constructor; }; }();

var _get = function get(object, property, receiver) { if (object === null) object = Function.prototype; var desc = Object.getOwnPropertyDescriptor(object, property); if (desc === undefined) { var parent = Object.getPrototypeOf(object); if (parent === null) { return undefined; } else { return get(parent, property, receiver); } } else if ("value" in desc) { return desc.value; } else { var getter = desc.get; if (getter === undefined) { return undefined; } return getter.call(receiver); } };

var _react = require('react');

var _react2 = _interopRequireDefault(_react);

var _reactDom = require('react-dom');

var _reactDom2 = _interopRequireDefault(_reactDom);

function _interopRequireDefault(obj) { return obj && obj.__esModule ? obj : { default: obj }; }

function _classCallCheck(instance, Constructor) { if (!(instance instanceof Constructor)) { throw new TypeError("Cannot call a class as a function"); } }

function _possibleConstructorReturn(self, call) { if (!self) { throw new ReferenceError("this hasn't been initialised - super() hasn't been called"); } return call && (typeof call === "object" || typeof call === "function") ? call : self; }

function _inherits(subClass, superClass) { if (typeof superClass !== "function" && superClass !== null) { throw new TypeError("Super expression must either be null or a function, not " + typeof superClass); } subClass.prototype = Object.create(superClass && superClass.prototype, { constructor: { value: subClass, enumerable: false, writable: true, configurable: true } }); if (superClass) Object.setPrototypeOf ? Object.setPrototypeOf(subClass, superClass) : subClass.__proto__ = superClass; }

var types = _react2.default.PropTypes;

function rootClick(e) {
  if (this.props.disableScrim || _reactDom2.default.findDOMNode(this).contains(e.target)) return;
  this.scrimClick(e);
}

var privates = new WeakMap();

module.exports = function (ParentClass) {
  var _class, _temp;

  return _temp = _class = function (_ParentClass) {
    _inherits(Scrim, _ParentClass);

    function Scrim(props, context) {
      _classCallCheck(this, Scrim);

      var _this = _possibleConstructorReturn(this, Object.getPrototypeOf(Scrim).call(this, props, context));

      privates.set(_this, rootClick.bind(_this));
      return _this;
    }

    _createClass(Scrim, [{
      key: 'scrimClick',
      value: function scrimClick() {}
    }, {
      key: 'componentDidMount',
      value: function componentDidMount() {
        var _get2;

        for (var _len = arguments.length, args = Array(_len), _key = 0; _key < _len; _key++) {
          args[_key] = arguments[_key];
        }

        if (_get(Object.getPrototypeOf(Scrim.prototype), 'componentDidMount', this)) (_get2 = _get(Object.getPrototypeOf(Scrim.prototype), 'componentDidMount', this)).call.apply(_get2, [this].concat(args));
        document.documentElement.addEventListener('click', privates.get(this));
      }
    }, {
      key: 'componentWillUnmount',
      value: function componentWillUnmount() {
        var _get3;

        for (var _len2 = arguments.length, args = Array(_len2), _key2 = 0; _key2 < _len2; _key2++) {
          args[_key2] = arguments[_key2];
        }

        if (_get(Object.getPrototypeOf(Scrim.prototype), 'componentWillUnmount', this)) (_get3 = _get(Object.getPrototypeOf(Scrim.prototype), 'componentWillUnmount', this)).call.apply(_get3, [this].concat(args));
        document.documentElement.removeEventListener('click', privates.get(this));
      }
    }]);

    return Scrim;
  }(ParentClass), _class.propTypes = {
    disableScrim: types.bool
  }, _temp;
};