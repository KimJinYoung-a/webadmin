'use strict';

angular.module('naldoApp')
    .config(function ($stateProvider) {
        $stateProvider
            .state('api', {
                abstract: true,
                parent: 'site'
            });
    });
