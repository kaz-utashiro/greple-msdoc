requires 'perl', '5.014';

requires 'App::Greple', '8.26';
requires 'App::optex::textconv';

on 'test' => sub {
    requires 'Test::More', '0.98';
};

