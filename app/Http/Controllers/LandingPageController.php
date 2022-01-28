<?php

namespace App\Http\Controllers;

use Illuminate\Http\Request;
use File;


class LandingPageController extends Controller
{
    /**
     * Create a new controller instance.
     *
     * @return void
     */
    public function index()
    {
          return File::get(public_path() . '/BUMDES Landing/index.html');   

    }
}
