package com.stefanini.apps;

/**
 * Hello world!
 *
 */
public class App {
	public static void main(String[] args) {
		try {
			SkillMapResultCollecter collecter = new SkillMapResultCollecter();
			collecter.collect(args.length >= 1 ? args[0] : null, args.length >= 2 ? args[1] : null);
		} catch (Exception e) {
			e.printStackTrace();
		}
	}
}
