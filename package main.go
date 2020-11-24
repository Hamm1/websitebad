package main

import (
	"fmt"
	"net/http"

	"github.com/gorilla/mux"
)

type application struct{}

func main() {

	app := &application{}

	srv := &http.Server{
		Handler: app.routes(),
		Addr:    ":80",
	}

	err := srv.ListenAndServe()
	if err != nil {
		fmt.Println(err)
	}

}
func (app *application) serveFilesavalon(w http.ResponseWriter, r *http.Request) {
	p := "." + r.URL.Path
	if p == "./" {
		p = "./avalon/index.html"
	}
	fmt.Println(p)
	http.ServeFile(w, r, p)
}
func (app *application) serveFilesgreeno(w http.ResponseWriter, r *http.Request) {
	p := "." + r.URL.Path
	if p == "./" {
		p = "./greeno/index.html"
	}
	fmt.Println(p)
	http.ServeFile(w, r, p)
}
func (app *application) serveFilesmini(w http.ResponseWriter, r *http.Request) {
	p := "." + r.URL.Path
	if p == "./" {
		p = "./mini/index.html"
	}
	fmt.Println(p)
	http.ServeFile(w, r, p)
}
func (app *application) routes() http.Handler {
	mux := mux.NewRouter()
	//mux.HandleFunc("/Action(ConfigurationId='{ConfigurationId}')/ConfigurationContent", app.GetConfiguration)
	mux.HandleFunc("/avalon/", app.serveFilesavalon)
	mux.HandleFunc("/greeno/", app.serveFilesgreeno)
	mux.HandleFunc("/mini/", app.serveFilesmini)

	return mux
}
